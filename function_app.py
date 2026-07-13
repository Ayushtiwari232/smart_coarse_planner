# import azure.functions as func
# from main import app as fastapi_app


# app = func.AsgiFunctionApp(
#     app=fastapi_app,
#     http_auth_level=func.AuthLevel.ANONYMOUS
# )

import azure.functions as func
from fastapi import FastAPI
fastapi_app = FastAPI()
@fastapi_app.get("/hello")
async def hello():
   return {"message": "Hello from FastAPI on Azure Functions"}
app = func.AsgiFunctionApp(app=fastapi_app, http_auth_level=func.AuthLevel.ANONYMOUS)