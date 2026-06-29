const DEFAULT_MODEL = "gpt-5.4-mini";

const SYSTEM_PROMPT = [
  "You are a Tradedoubler publisher reporting assistant for QBR slides.",
  "",
  "Goal:",
  "Produce advertiser-style publisher QBR narrative grounded only in supplied data.",
  "",
  "Rules:",
  "- Write narrative text in the requested language from the user payload.",
  "- Keep numeric values, dates, percentages, table values, and currency symbols exactly as supplied.",
  "- Keep insights concise and factual.",
  "- No invented causes, benchmarks, or recommendations unsupported by data.",
  "- If data is missing, state it explicitly.",
  "",
  "Return strict JSON only, with keys: reportingPeriod, kpiHighlights, programLevelAnalysis, moversAndShakers, risksAndDependencies."
].join("\n");

function asText(value, fallback = "") {
  if (value === null || value === undefined) return fallback;
  const text = String(value).trim();
  return text || fallback;
}

function redactSensitive(value) {
  const sensitiveKey = /token|authorization|password|secret|api[-_]?key|cookie|credential/i;
  if (Array.isArray(value)) return value.map(redactSensitive);
  if (!value || typeof value !== "object") {
    if (typeof value === "string") {
      return value
        .replace(/bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [redacted]")
        .replace(/sk-[A-Za-z0-9_-]+/g, "sk-[redacted]");
    }
    return value;
  }

  return Object.fromEntries(Object.entries(value).map(([key, item]) => {
    if (sensitiveKey.test(key)) return [key, item ? "[redacted]" : item];
    return [key, redactSensitive(item)];
  }));
}

function renderUserPrompt(context = {}) {
  const payload = context.payload || {};
  const focus = asText(payload.qbrFocus, "General performance review");
  const focusDetail = asText(payload.qbrFocusDetail, "N/A");
  const languageName = asText(payload.languageName, "English");
  const dataForAI = asText(context.dataForAI, "{}");

  return [
    dataForAI,
    "",
    "You are preparing the Publisher Performance section for a client QBR.",
    "",
    "PRIMARY QBR FOCUS:",
    `- Focus theme: ${focus}`,
    `- Focus detail: ${focusDetail}`,
    "",
    "MODE: CONTEXT-RICH REPORTING",
    "- Keep output factual, concise, and presentation-ready.",
    "- Provide concise, data-backed implications; recommendations are allowed only when explicitly supported by provided data.",
    "- Do not infer causes unless directly evidenced in the provided data.",
    "- If cause is unclear, state: \"Driver not confirmed from available data.\"",
    "",
    "OUTPUT STRUCTURE:",
    "Return strict JSON with this shape:",
    "{",
    "  \"reportingPeriod\": [\"...\"],",
    "  \"kpiHighlights\": [\"...\"],",
    "  \"programLevelAnalysis\": [{\"title\":\"...\",\"description\":\"...\"}],",
    "  \"moversAndShakers\": [\"...\"],",
    "  \"risksAndDependencies\": [{\"title\":\"...\",\"description\":\"...\"}]",
    "}",
    "",
    "LANGUAGE REQUIREMENT:",
    `Write all narrative text and bullets in ${languageName}.`,
    "",
    "NUMBER & DATE REQUIREMENT:",
    "- Use UK number formatting only: 1,234.56",
    "- Use currency symbol before number (e.g. £1,899.64)",
    "- Use % symbol for rates",
    "- Keep numeric values and table structures exactly as provided.",
    "",
    "API SOURCE SCOPE REQUIREMENT:",
    "- For Executive Summary bottom analysis and KPI Highlights, use full API program scope from tables.allProgramsApiScope plus KPI totals.",
    "- Treat programLevelBreakdown as a display subset only.",
    "- Do not limit narrative analysis to any 12-row subset when allProgramsApiScope is provided."
  ].join("\n");
}

function extractResponseContent(response) {
  if (typeof response === "string") return response;
  if (!response || typeof response !== "object") return "";
  if (typeof response.content === "string") return response.content;
  if (typeof response.output_text === "string") return response.output_text;
  if (typeof response.text === "string") return response.text;
  if (response.message && typeof response.message.content === "string") return response.message.content;
  return "";
}

function parseAgentJson(content) {
  const raw = asText(content);
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const jsonText = fenced ? fenced[1].trim() : raw;
  try {
    return JSON.parse(jsonText);
  } catch (error) {
    throw new Error("Agent output must be valid JSON.");
  }
}

function validateStringArray(value, key) {
  if (!Array.isArray(value)) throw new Error(`Agent output field ${key} must be an array.`);
  for (const item of value) {
    if (typeof item !== "string") throw new Error(`Agent output field ${key} must contain strings.`);
  }
}

function validateTitledRows(value, key) {
  if (!Array.isArray(value)) throw new Error(`Agent output field ${key} must be an array.`);
  for (const item of value) {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      throw new Error(`Agent output field ${key} must contain objects.`);
    }
    if (typeof item.title !== "string" || typeof item.description !== "string") {
      throw new Error(`Agent output field ${key} rows need title and description strings.`);
    }
  }
}

function validateAgentOutput(output) {
  if (!output || typeof output !== "object" || Array.isArray(output)) {
    throw new Error("Agent output must be a JSON object.");
  }
  validateStringArray(output.reportingPeriod, "reportingPeriod");
  validateStringArray(output.kpiHighlights, "kpiHighlights");
  validateTitledRows(output.programLevelAnalysis, "programLevelAnalysis");
  validateStringArray(output.moversAndShakers, "moversAndShakers");
  validateTitledRows(output.risksAndDependencies, "risksAndDependencies");
  return output;
}

function markdownList(lines) {
  return (lines || []).map((line) => `- ${line}`).join("\n");
}

function markdownRows(rows) {
  return (rows || []).map((row) => `- ${row.title}: ${row.description}`).join("\n");
}

function agentJsonToMarkdown(output) {
  return [
    "## Publisher Performance",
    "",
    "### Reporting Period",
    markdownList(output.reportingPeriod),
    "",
    "### KPI Highlights (Detailed)",
    markdownList(output.kpiHighlights),
    "",
    "### Program Level Analysis (Top Performers)",
    markdownRows(output.programLevelAnalysis),
    "",
    "### Movers & Shakers",
    markdownList(output.moversAndShakers),
    "",
    "### Risks and Dependencies",
    markdownRows(output.risksAndDependencies)
  ].join("\n").trim();
}

function firstRows(rows, limit = 3) {
  return Array.isArray(rows) ? rows.slice(0, limit) : [];
}

function describeRow(row) {
  if (!row || typeof row !== "object" || Array.isArray(row)) return asText(row, "N/A");
  const preferred = [
    "programName",
    "Program Name",
    "name",
    "Program",
    "programId",
    "Program ID",
    "publisherCommission",
    "Publisher Commission",
    "Rows",
    "Value"
  ];
  const parts = [];
  for (const key of preferred) {
    if (row[key] !== undefined && row[key] !== null && asText(row[key])) {
      parts.push(`${key}: ${asText(row[key])}`);
    }
  }
  if (parts.length) return parts.slice(0, 4).join(", ");
  return Object.entries(row).slice(0, 4).map(([key, value]) => `${key}: ${asText(value, "-")}`).join(", ");
}

function parseDataForAI(dataForAI) {
  const parsed = parseAgentJson(asText(dataForAI, "{}"));
  return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
}

function createDeterministicAgentOutput(context = {}) {
  const payload = context.payload || {};
  const data = parseDataForAI(context.dataForAI);
  const tables = data.tables || context.tables || {};
  const diagnostics = data.diagnostics || context.diagnostics || {};
  const reportingPeriod = asText(payload.reportingPeriod || data.payload?.reportingPeriod, "Not specified");
  const comparisonPeriod = asText(payload.comparisonPeriod || data.payload?.comparisonPeriod, "Not specified");
  const currentRows = diagnostics.currentRows ?? firstRows(tables.programLevelBreakdown, 9999).length;
  const previousRows = diagnostics.previousRows ?? "N/A";
  const competitorPublishers = diagnostics.competitorPublishers ?? 0;
  const topPrograms = firstRows(tables.programLevelBreakdown || tables.allProgramsApiScope, 3);
  const riskRows = firstRows(tables.riskDependenciesTable, 5);

  return validateAgentOutput({
    reportingPeriod: [
      `Current period ${reportingPeriod}; YoY comparison period ${comparisonPeriod}.`
    ],
    kpiHighlights: [
      `Workflow processed ${currentRows} current publisher rows and ${previousRows} previous publisher rows from the supplied API data.`,
      `Competitor context includes ${competitorPublishers} comparison publisher(s).`,
      `QBR focus: ${asText(payload.qbrFocus, "General performance review")}.`
    ],
    programLevelAnalysis: topPrograms.length
      ? topPrograms.map((row, index) => ({
          title: `Program signal ${index + 1}`,
          description: describeRow(row)
        }))
      : [{ title: "Program data availability", description: "No program-level rows were available in the computed workflow tables." }],
    moversAndShakers: firstRows(tables.moversShakersCommissionChart || tables.programLevelBreakdown, 5)
      .map((row) => describeRow(row)),
    risksAndDependencies: riskRows.length
      ? riskRows.map((row) => ({
          title: asText(row.Issue || row.title || row.Metric, "Risk signal"),
          description: asText(row.Analysis || row.description || row.Value || describeRow(row), "Risk detail not supplied.")
        }))
      : []
  });
}

function validateToolInput(schema, input, path = "input") {
  if (!schema) return;
  if (schema.type === "object") {
    if (!input || typeof input !== "object" || Array.isArray(input)) {
      throw new Error(`${path} must be an object.`);
    }
    for (const key of schema.required || []) {
      if (input[key] === undefined) throw new Error(`${path}.${key} is required.`);
    }
    for (const [key, property] of Object.entries(schema.properties || {})) {
      if (input[key] === undefined) continue;
      validateToolInput(property, input[key], `${path}.${key}`);
    }
    return;
  }
  if (schema.type && typeof input !== schema.type) {
    throw new Error(`${path} must be ${schema.type}.`);
  }
}

function describeTools(tools) {
  return Object.entries(tools || {}).map(([name, tool]) => ({
    name,
    description: tool.description || "",
    schema: tool.schema || {}
  }));
}

function createDefaultModelClient(fetchImpl = globalThis.fetch) {
  return {
    async complete(request) {
      const apiKey = process.env.OPENAI_API_KEY;
      if (!apiKey) {
        throw new Error("OPENAI_API_KEY is required for live local agent runs. Tests should inject modelClient.");
      }
      const response = await fetchImpl("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${apiKey}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: request.model,
          instructions: request.system,
          input: request.user
        })
      });
      const text = await response.text();
      let body;
      try {
        body = text ? JSON.parse(text) : {};
      } catch {
        body = { text };
      }
      if (!response.ok) {
        throw new Error(`OpenAI model call failed: HTTP ${response.status}`);
      }
      return {
        content: body.output_text || body.text || text
      };
    }
  };
}

function createPublisherQbrAgent(options = {}) {
  const model = options.model || DEFAULT_MODEL;
  const shouldUseDeterministic = options.deterministic === true
    || (!options.modelClient && !process.env.OPENAI_API_KEY)
    || String(process.env.PUBLISHER_QBR_AGENT_MODE || "").toLowerCase() === "deterministic";
  const modelClient = shouldUseDeterministic ? null : (options.modelClient || createDefaultModelClient(options.fetch));
  const tools = options.tools || {};
  const maxIterations = Math.max(1, Number(options.maxIterations || 4));
  const logger = options.logger || console;
  const fallbackOnInvalidOutput = options.fallbackOnInvalidOutput !== false;

  return {
    async run(context = {}) {
      if (shouldUseDeterministic) {
        logger.warn?.("publisher_qbr_agent_deterministic_fallback", redactSensitive({
          reason: "OPENAI_API_KEY not set or deterministic mode requested",
          payload: context.payload
        }));
        const json = createDeterministicAgentOutput(context);
        return {
          json,
          markdown: agentJsonToMarkdown(json),
          toolResults: [],
          deterministic: true
        };
      }

      const system = SYSTEM_PROMPT;
      const user = renderUserPrompt(context);
      const toolResults = [];

      for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
        logger.info?.("publisher_qbr_agent_model_call", redactSensitive({
          model,
          iteration,
          payload: context.payload,
          toolResultCount: toolResults.length
        }));
        const response = await modelClient.complete({
          model,
          system,
          user,
          tools: describeTools(tools),
          toolResults
        }, { iteration, toolResults });

        const toolCalls = response?.toolCalls || response?.tool_calls || (response?.toolCall ? [response.toolCall] : []);
        if (toolCalls.length) {
          if (iteration >= maxIterations) throw new Error("Agent reached maximum agent iterations before final output.");
          for (const call of toolCalls) {
            const name = call.name || call.toolName || call.function?.name;
            const input = call.input || call.arguments || call.function?.arguments || {};
            const tool = tools[name];
            if (!tool) throw new Error(`Tool ${name} is not allowlisted.`);
            const parsedInput = typeof input === "string" ? JSON.parse(input) : input;
            validateToolInput(tool.schema, parsedInput);
            logger.info?.("publisher_qbr_agent_tool_call", redactSensitive({ name, input: parsedInput }));
            const result = await tool.execute(parsedInput, context);
            toolResults.push({ name, input: parsedInput, result });
          }
          continue;
        }

        try {
          const parsed = validateAgentOutput(parseAgentJson(extractResponseContent(response)));
          return {
            json: parsed,
            markdown: agentJsonToMarkdown(parsed),
            toolResults
          };
        } catch (error) {
          if (!fallbackOnInvalidOutput) throw error;
          logger.warn?.("publisher_qbr_agent_invalid_output_fallback", redactSensitive({
            error: error instanceof Error ? error.message : String(error),
            model,
            iteration
          }));
          const fallbackJson = createDeterministicAgentOutput(context);
          return {
            json: fallbackJson,
            markdown: agentJsonToMarkdown(fallbackJson),
            toolResults,
            deterministic: true,
            fallbackReason: error instanceof Error ? error.message : String(error)
          };
        }
      }

      throw new Error("Agent reached maximum agent iterations before final output.");
    }
  };
}

module.exports = {
  SYSTEM_PROMPT,
  DEFAULT_MODEL,
  createPublisherQbrAgent,
  parseAgentJson,
  validateAgentOutput,
  agentJsonToMarkdown,
  redactSensitive,
  renderUserPrompt
};
