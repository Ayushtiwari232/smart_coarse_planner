import os

from dotenv import load_dotenv
from langchain_openai import AzureChatOpenAI

load_dotenv()


def get_llm() -> AzureChatOpenAI:
    """Return an AzureChatOpenAI instance configured for GPT-5 Nano."""
    return AzureChatOpenAI(
        azure_endpoint=os.getenv("GPT5_NANO_ENDPOINT"),
        api_key=os.getenv("GPT5_NANO_API_KEY"),
        api_version=os.getenv("GPT5_NANO_API_VERSION"),
        azure_deployment=os.getenv("GPT5_NANO_MODEL_NAME"),
    )
