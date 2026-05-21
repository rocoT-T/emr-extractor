# 导入FastAPI类
# from fastapi import FastAPI

# # 创建FastAPI应用实例
# # title是API的名称，会显示在自动生成的文档里
# app = FastAPI(title="电子病历信息提取器API", version="0.2.0")

# # 定义根路径的GET接口
# # @app.get("/") 是装饰器，表示这个函数处理 GET 请求，路径是 "/"
# @app.get("/")
# def root():
#     """根接口，返回API运行状态"""
#     return {
#         "message": "EMR Extractor API is running",
#         "version": "0.2.0",
#         "status": "success"
#     }

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from docx import Document
import re
from datetime import datetime
import os
import tempfile

# --------------------------
# 导入FastAPI类
from fastapi import FastAPI

# 创建FastAPI应用实例
# title是API的名称，会显示在自动生成的文档里
app = FastAPI(title="电子病历信息提取器API", version="0.2.0")

# 定义根路径的GET接口
# @app.get("/") 是装饰器，表示这个函数处理 GET 请求，路径是 "/"
@app.get("/")
def root():
    """根接口，返回API运行状态"""
    return {
        "message": "EMR Extractor API is running",
        "version": "0.2.0",
        "status": "success"
    }
# --------------------------

app = FastAPI(title="电子病历信息提取器API", version="0.2.0")

@app.get("/")
def root():
    return {
        "message": "EMR Extractor API is running",
        "version": "0.2.0",
        "status": "success"
    }

@app.post("/extract", summary="提取病历信息")
async def extract_emr(file: UploadFile = File(..., description="上传Word格式的病历文件(.docx)")):
    # 就是上面写的那个接口代码
    pass