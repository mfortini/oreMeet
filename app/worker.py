import pandas as pd
import numpy as np
import datetime
import logging
import io

def roundTime(t, step, direction):
    fracSteps = t / (step * 1.)

    if direction == "round":
        intSteps = int(round(fracSteps))
    elif direction == "up":
        intSteps = int(np.ceil(fracSteps))
    elif direction == "down":
        intSteps = int(np.floor(fracSteps))

    return step * intSteps

def formatSeconds(t):
    return "%02d:%02d:%02d" % (t//3600, (t%3600)//60, t%60)

def meetReport (csvFile, meetingCode, _startTime, _endTime, startStep, startRoundDir, midThr, midStep, midRoundDir, endStep, endRoundDir):

    meetingStartTime=pd.to_datetime(_startTime).tz_localize(None)
    meetingEndTime=pd.to_datetime(_endTime).tz_localize(None)


    meetData = pd.read_csv(csvFile)
    meetingData = meetData[meetData["Codice riunione"]==meetingCode]
    meetingData["Data"]=pd.to_datetime(meetingData["Data"]).dt.tz_localize(None)
    meetingData["endTime"]=pd.to_datetime(meetingData["Data"]).dt.tz_localize(None)
    meetingData["startTime"]=list(map(lambda x:x[1]["endTime"]-pd.to_timedelta(x[1]["Durata"],'s'),meetingData.iterrows()))

    meetingData.loc[:,["endTime","startTime"]]=meetingData[["endTime","startTime"]].clip(lower=meetingStartTime,upper=meetingEndTime)

    participantsTime=meetingData.groupby(["Identificatore partecipante","Nome evento"])

    results=[]

    for name,g in participantsTime:
        participantName = name[0]
        participantRes={
                "Nome": participantName,
                "Assenza intermedia sogliata":0,
                "Assenza intermedia arrotondata":0,
                "Assenza intermedia":0,
                "Presente netto":0,
                }
        beginning=True
        logging.debug("participantName {}".format(participantName))
        intervals = []
        for itt,tt in g.sort_values(by="startTime").iterrows():
            if len(intervals) == 0 or tt["startTime"] > intervals[-1]["endTime"]:
                intervals.append(tt[["startTime","endTime"]])
            else:
                logging.warning("overlapping intervals {} and {}".format(intervals[-1],tt[["startTime", "endTime"]]))
                intervals[-1]["endTime"]=tt["endTime"]

        curTime=meetingStartTime
        for iv in intervals:
            logging.debug("curTime {} startTime {} endTime {}".format(curTime, iv["startTime"],iv["endTime"]))
            diffTime=(iv["startTime"]-curTime).seconds
            logging.debug("diffTime {}".format(diffTime))
            if beginning:
                beginning=False
                participantRes["Orario entrata"]=iv["startTime"]
                participantRes["Assenza inizio"]=diffTime
                participantRes["Assenza inizio arrotondata"]=roundTime(diffTime, startStep, startRoundDir)
            else:
                participantRes["Assenza intermedia sogliata"] += diffTime if diffTime >= midThr else 0
                participantRes["Assenza intermedia"] += diffTime
            participantRes["Presente netto"] += (iv["endTime"]-iv["startTime"]).seconds
            curTime=iv["endTime"]
        diffTime=meetingEndTime-curTime
        participantRes["Assenza fine"]=diffTime.seconds
        participantRes["Assenza fine arrotondata"]=roundTime(diffTime.seconds, endStep, endRoundDir)
        participantRes["Orario uscita"]=curTime
        participantRes["Assenza intermedia arrotondata"] = roundTime(participantRes["Assenza intermedia sogliata"], midStep, midRoundDir)
        participantRes["Presente complessivo arrotondato"]=(meetingEndTime - meetingStartTime).seconds - participantRes["Assenza inizio arrotondata"] - participantRes["Assenza fine arrotondata"] - participantRes["Assenza intermedia arrotondata"]
        
        results.append(participantRes)

    output=io.StringIO()
    dfResults=pd.DataFrame(results)

    for col in ["Assenza inizio arrotondata", "Assenza intermedia arrotondata", "Assenza fine arrotondata", "Presente complessivo arrotondato"]:
        dfResults[col] = list(map(lambda x:formatSeconds(x),dfResults[col]))

    dfResults.to_csv(output)

    return output


if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.DEBUG)
    logging.info("Open {}".format([sys.argv[1]]))
    with open(sys.argv[1]) as f:
        report = meetReport (f, sys.argv[2], sys.argv[3], sys.argv[4], int(sys.argv[5]), sys.argv[6], int(sys.argv[7]), sys.argv[8], int(sys.argv[9]), sys.argv[10])
        print(pd.DataFrame(report))
