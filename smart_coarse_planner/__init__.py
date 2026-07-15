import azure.functions as func
import logging
from typing import Annotated, Literal, Sequence
from typing_extensions import TypedDict
import asyncio
import json
import re
import traceback
from .tracing import tracer,root_span
from .utils import parse_json_from_markdown,convert_links_to_markdown,get_link,get_actions,extract_links,extract_answer_source_json
from langchain_core.messages import BaseMessage, HumanMessage
from langchain_core.messages import HumanMessage, AIMessage, ToolMessage
from langchain_core.tools import tool
from dotenv import load_dotenv
from langchain_core.runnables import RunnableConfig
from langgraph.checkpoint.memory import InMemorySaver
import os
from langchain_openai import AzureOpenAIEmbeddings
from langchain_openai import AzureChatOpenAI
from langchain_community.vectorstores.azuresearch import AzureSearch
from .prompts import generate_prompt_template
import aiosqlite
from langgraph.checkpoint.sqlite.aio import AsyncSqliteSaver
from .agent import AzureHrAgent,treasury_retriever_tool,snow_retriever_tool
from opencensus.ext.azure.log_exporter import AzureLogHandler


import azure.functions as func
import json
import logging


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Health endpoint called.")

    response = {
        "status": "ok",
        "message": "Azure Function is running"
    }

    return func.HttpResponse(
        body=json.dumps(response),
        mimetype="application/json",
        status_code=200
    )