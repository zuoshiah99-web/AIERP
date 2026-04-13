import os
import pandas as pd
import pyodbc
from fastapi import FastAPI, Request, Form
from fastapi.templating import Jinja2Templates
from fastapi.responses import FileResponse
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()
app = FastAPI()
templates = Jinja2Templates(directory="templates")

# 初始化 Gemini
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-1.5-flash')

# 資料庫連線字串
conn_str = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={os.getenv('DB_SERVER')};"
    f"DATABASE={os.getenv('DB_DATABASE')};"
    f"UID={os.getenv('DB_USER')};"
    f"PWD={os.getenv('DB_PWD')}"
)

def get_schema():
    # 獲取視圖(View)的結構資訊，注入給 Gemini
    sql = """
    SELECT TABLE_NAME, COLUMN_NAME 
    FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME IN ('科目餘額表', '採購明細表')
    """
    conn = pyodbc.connect(conn_str)
    df = pd.read_sql(sql, conn)
    conn.close()
    return df.to_string()

@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/query")
async def query(request: Request, prompt: str = Form(...)):
    schema = get_schema()
    
    # RAG 注入：將 Schema 餵給 Gemini
    ai_prompt = f"""
    你是一個 SQL Server 專家。請根據以下 Schema 資訊：
    {schema}
    
    將使用者的需求轉化為 SQL 查詢語句。
    規則：
    1. 只回傳 SQL 程式碼，不要有 Markdown 格式。
    2. 使用繁體中文標籤。
    3. 需求內容：{prompt}
    """
    
    response = model.generate_content(ai_prompt)
    generated_sql = response.text.strip()
    
    # 執行 SQL
    conn = pyodbc.connect(conn_str)
    try:
        df = pd.read_sql(generated_sql, conn)
        # 數字欄位自動加總
        totals = df.select_dtypes(include=['number']).sum().to_dict()
        # 存成 Excel 供下載
        df.to_excel("temp_result.xlsx", index=False)
        result_html = df.to_html(classes="table table-bordered table-striped", index=False)
        error = None
    except Exception as e:
        result_html = ""
        totals = {}
        error = str(e)
    finally:
        conn.close()
        
    return templates.TemplateResponse("result.html", {
        "request": request, 
        "sql": generated_sql, 
        "table": result_html, 
        "totals": totals,
        "error": error
    })

@app.get("/export")
async def export():
    return FileResponse("temp_result.xlsx", filename="ERP_查詢結果.xlsx")