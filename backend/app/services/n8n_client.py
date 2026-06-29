import httpx
from typing import Any


class N8nClient:
    """Legacy app-hub compatibility client for retired workflow webhooks."""

    def __init__(self, timeout: float = 360.0):
        self.timeout = timeout

    async def call_webhook(
        self,
        webhook_url: str,
        message: str,
        thread_id: str | None = None,
        extra_data: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """
        Call a legacy workflow webhook with a message.

        Args:
            webhook_url: The legacy webhook URL
            message: The user's message
            thread_id: Optional thread ID for conversation context
            extra_data: Optional additional data to send

        Returns:
            The response from the legacy workflow
        """
        payload = {
            "message": message,
            "thread_id": thread_id,
        }

        if extra_data:
            payload.update(extra_data)

        async with httpx.AsyncClient(trust_env=False) as client:
            response = await client.post(
                webhook_url,
                json=payload,
                timeout=self.timeout,
            )
            response.raise_for_status()
            try:
                return response.json()
            except ValueError:
                # Some legacy webhooks return plain text or an empty body on success.
                text = response.text.strip()
                return {"response": text or "Webhook triggered successfully."}


# Singleton instance
n8n_client = N8nClient()


