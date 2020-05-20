from fastapi import FastAPI, File, Form, UploadFile, Request, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import StreamingResponse

from worker import meetReport

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")


@app.get("/")
async def read_item(request: Request):
    return templates.TemplateResponse("inputForm.html",{"request": request})

@app.get("/items/{id}")
async def read_item(request: Request, id: str):
    return templates.TemplateResponse("item.html", {"request": request, "id": id})



@app.post("/meetReport/")
async def create_file(
    file: UploadFile = File(...),
    calFile: UploadFile = File(None),
    meetingCode: str = Form(...),
    startTime: str = Form(...),
    endTime: str = Form(...),
    startStep: int = Form(...),
    startRoundDir: str = Form(...),
    midThr: int = Form(...),
    midStep: int = Form(...),
    midRoundDir: str = Form(...),
    endStep: int = Form(...),
    endRoundDir: str = Form(...),
):

    print(calFile)
    result=meetReport(file.filename,file.file, calFile.file if calFile else None, meetingCode, startTime, endTime, startStep, startRoundDir, midThr, midStep, midRoundDir, endStep, endRoundDir)
    result.seek(0)
    return StreamingResponse(
            result,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            #headers={'Content-Disposition': 'attachment;filename="pippo.csv"'}
        )

