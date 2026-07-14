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

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')
    
    logger = logging.getLogger(__name__)
    logger.setLevel("INFO")
    logger.addHandler(AzureLogHandler(connection_string="InstrumentationKey=f3f1f85f-a4ab-4514-b91e-223ee0f677aa;IngestionEndpoint=https://centralus-2.in.applicationinsights.azure.com/;LiveEndpoint=https://centralus.livediagnostics.monitor.azure.com/;ApplicationId=a87e1eee-3d81-4b10-b8a2-dd556d3e360a"))


    try:
        # Get request body..
        req_text = req.get_body()
        logger.info(f"Raw request body: {req_text}")

        req_body = req.get_json()
        question = req_body.get('question')
        sessionID = req_body.get('sessionID')

        logger.info(f"Question: {question}")
        logger.info(f"Session ID: {sessionID}")
        logger.info("new version")

        if sessionID and question:
            try:
                res = asyncio.run(chat(sessionID, question,logger))
                if res:
                # trace_details = traceback.format_exc()
                    tool_calls = res['messages'][1].additional_kwargs.get('tool_calls')
                    if tool_calls:
                        try:
                           response=parse_json_from_markdown(res['messages'][-1].content)
                           markdown_answer = convert_links_to_markdown(response['Answer'])
                        #    markdown_answer=response['Answer']
                           result=get_link(response['Source'])
                           actions = []
                           for i in result:
                                a = {
                                            "type": "Action.OpenUrl",
                                            "title": i[0],
                                            "url": i[1]
                                        }
                                actions.append(a)
                           if actions: 
                                links_adap =get_actions(actions)
                                output={"answer":markdown_answer,"source":str(response["Source"]),"question":question,"KBID":response['Source'],"link_adap":links_adap,"isAnswered":True}
                           else:
                                output={"answer":markdown_answer,"source":"no data","question":question,"KBID":"No data","isAnswered":False}
                           return func.HttpResponse(json.dumps(output))
                        except:
                            # trace_details = traceback.format_exc()
                            markdown_answer = convert_links_to_markdown(res['messages'][-1].content)
                            # markdown_answer=res['messages'][-1].content
                            source_links=extract_links(res['messages'][-1].content)
                            if source_links:
                                actions = []
                                for i,j in enumerate(source_links):
                                    
                                    a = {
                                                            "type": "Action.OpenUrl",
                                                            "title": "Source "+str(i+1),
                                                            "url":j.lstrip('(').rstrip(').')
                                                        }
                                    actions.append(a)
                                if actions:
                                    links_adap =get_actions(actions)
                                    output={"answer":markdown_answer,"source":"No data","question":question,"KBID":"No data","link_adap":links_adap,"isAnswered":True}                            
                                else:
                                    output={"answer":markdown_answer,"source":"No data","question":question,"KBID":"No data","isAnswered":False}
                            else:
                                result=extract_answer_source_json(markdown_answer)
                                if result.get("Answer"):
                                    markdown_answer = convert_links_to_markdown(result['Answer'])
                                    # markdown_answer=result['Answer']
                                    if result.get("Source"): 
                                        Source_list=get_link(result['Source'])
                                        actions = []
                                        for i in Source_list:
                                                a = {
                                                            "type": "Action.OpenUrl",
                                                            "title": i[0],
                                                            "url": i[1]
                                                        }
                                                actions.append(a)
                                        if actions:
                                            links_adap =get_actions(actions)
                                            output={"answer":markdown_answer,"source":str(result["Source"]),"question":question,"KBID":result['Source'],"link_adap":links_adap,"isAnswered":True}
                                        else:
                                            output={"answer":markdown_answer,"source":"No data","question":question,"KBID":"No data","isAnswered":True}
                                    else:
                                        output={"answer":markdown_answer,"source":"No data","question":question,"KBID":"No data","isAnswered":False}
                                else:
                                    output = {"answer":markdown_answer,"source": "No data", "question": question, "KBID":"No data","isAnswered":False}
                            return func.HttpResponse(json.dumps(output))
                    else:
                        default_response=" I am BlueBot, an advanced virtual assistant designed to help users with IT-related questions . My purpose is to provide accurate and detailed information to assist you with your IT inquiries."
                        output = {"answer":default_response,"source": "No data", "question": question, "KBID":"No data","isAnswered":False}
                        
                        return func.HttpResponse(json.dumps(output))
                else:
                    default_response="I am BlueBot, an advanced virtual assistant designed to help users with IT-related questions . My purpose is to provide accurate and detailed information to assist you with your IT inquiries."
                    output = {"answer":default_response,"source": "No data", "question": question, "KBID":"No data","isAnswered":False}
                    return func.HttpResponse(json.dumps(output))
            except Exception as e:
                trace_details = traceback.format_exc()
                logger.error(f"Error during chat function execution: {str(e)}")
                return func.HttpResponse(f"Error processing chat response. {str(e)}\nTraceback:\n{trace_details}", status_code=500)
        else:
            return func.HttpResponse(
                "Please provide both 'sessionID' and 'question' in the request body.",
                status_code=400
            )

    except Exception as e:
        logger.error(f"Failed to process request: {str(e)}")
        return func.HttpResponse(f"Failed to process request due to an internal error {str(e)}", status_code=500)

    
async def chat(session_id,question,logger):
    tools = [treasury_retriever_tool,snow_retriever_tool]
    # Create an instance of AzureSqlAgent with the checkpointer  
    sql_agent = AzureHrAgent(tools)  
    #from langchain_core.runnables import RunnableConfig
    config = RunnableConfig(recursion_limit=5)
    try:
        result = await sql_agent.graph.ainvoke(  
                    {"messages": [HumanMessage(content=question)]},  
                    config=config ,
                    
                )  
    except:
        result=None
    return result
