from dotenv import load_dotenv
from langchain_openai import AzureChatOpenAI
import os

load_dotenv()

def get_llm_from_env(prefix: str = "GPT5_NANO", temperature: float = 1) -> AzureChatOpenAI:
    """
    Create an AzureChatOpenAI client using a specific set of environment variables,
    allowing support for multiple models.

    Args:
        prefix (str): A prefix for the environment variables (e.g., "GPT4O", "GPT35").
        temperature (float): The temperature setting for the model.

    Required ENV variables (with the given prefix):
        {prefix}_API_KEY
        {prefix}_ENDPOINT
        {prefix}_MODEL
        {prefix}_API_VERSION

    Example:
        os.environ["GPT4O_API_KEY"] = "sk-..."
        os.environ["GPT4O_ENDPOINT"] = "https://...openai.azure.com/"
        os.environ["GPT4O_MODEL"] = "gpt-4o"
        os.environ["GPT4O_API_VERSION"] = "2024-05-01"

        llm = get_llm_from_env("GPT4O")
    """
    api_key = os.getenv(f"{prefix}_API_KEY")
    endpoint = os.getenv(f"{prefix}_ENDPOINT")
    model = os.getenv(f"{prefix}_MODEL")
    version = os.getenv(f"{prefix}_API_VERSION")

    if not all([api_key, endpoint, model, version]):
        raise ValueError(f"Missing one or more environment variables for prefix: {prefix}")

    return AzureChatOpenAI(
        openai_api_key=api_key,
        azure_endpoint=endpoint,
        deployment_name=model,
        openai_api_version=version,
        openai_api_type="azure",
    )