import pandas as pd
import numpy as np
import re
import datetime
import logging
import io
import ics
import locale

TZ="Europe/Rome"


tzMap={
        "ACDT": "+1030",
        "ACST": "+0930",
        "ADT": "-0300",
        "AEDT": "+1100",
        "AEST": "+1000",
        "AHDT": "-0900",
        "AHST": "-1000",
        "AST": "-0400",
        "AT": "-0200",
        "AWDT": "+0900",
        "AWST": "+0800",
        "BAT": "+0300",
        "BDST": "+0200",
        "BET": "-1100",
        "BST": "-0300",
        "BT": "+0300",
        "BZT2": "-0300",
        "CADT": "+1030",
        "CAST": "+0930",
        "CAT": "-1000",
        "CCT": "+0800",
        "CDT": "-0500",
        "CED": "+0200",
        "CET": "+0100",
        "CEST": "+0200",
        "CST": "-0600",
        "EAST": "+1000",
        "EDT": "-0400",
        "EED": "+0300",
        "EET": "+0200",
        "EEST": "+0300",
        "EST": "-0500",
        "FST": "+0200",
        "FWT": "+0100",
        "GMT": "GMT",
        "GST": "+1000",
        "HDT": "-0900",
        "HST": "-1000",
        "IDLE": "+1200",
        "IDLW": "-1200",
        "IST": "+0530",
        "IT": "+0330",
        "JST": "+0900",
        "JT": "+0700",
        "MDT": "-0600",
        "MED": "+0200",
        "MET": "+0100",
        "MEST": "+0200",
        "MEWT": "+0100",
        "MST": "-0700",
"MT": "+0800",
"NDT": "-0230",
"NFT": "-0330",
"NT": "-1100",
"NST": "+0630",
"NZ": "+1100",
"NZST": "+1200",
"NZDT": "+1300",
"NZT": "+1200",
"PDT": "-0700",
"PST": "-0800",
"ROK": "+0900",
"SAD": "+1000",
"SAST": "+0900",
"SAT": "+0900",
"SDT": "+1000",
"SST": "+0200",
"SWT": "+0100",
"USZ3": "+0400",
"USZ4": "+0500",
"USZ5": "+0600",
"USZ6": "+0700",
"UT": "-0000",
"UTC": "-0000",
"UZ10": "+1100",
"WAT": "-0100",
"WET": "-0000",
"WST": "+0800",
"YDT": "-0800",
"YST": "-0900",
"ZP4": "+0400",
"ZP5": "+0500",
"ZP6": "+0600"
}

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

def getEventMeetingCode(e):
    meetUrl=re.search(r"meet.google.com/(\w\w\w)-(\w\w\w\w)-(\w\w\w)", e.description)
    return (''.join(meetUrl.groups())).upper() if meetUrl else None

            

def convertDate(d):
    tzMatch = re.match(r'\s*\d+\s+\w+\s+\d{4},\s+\d+:\d+:\d+\s+(\S+)', d)
    if tzMatch:
        tz=tzMatch.group(1)
        try:
            tzNew = tzMap[tz]
            d=d.replace(tz,tzNew)
        except:
            pass

    locale.setlocale(locale.LC_ALL, "it_IT.utf8")
    try:
        return pd.to_datetime(d).tz_convert(TZ).tz_localize(None)
    except:
        pass

    try:
        return pd.to_datetime(d).tz_localize(None)
    except:
        pass

    locale.setlocale(locale.LC_ALL, "it_IT.utf8")
    formatStr="%d %b %Y, %H:%M:%S %Z"
    try:
        return pd.to_datetime(d, format=formatStr).tz_convert(TZ).tz_localize(None)
    except:
        pass

    formatStr="%d %b %Y, %H:%M:%S %z"
    try:
        return pd.to_datetime(d, format=formatStr).tz_convert(TZ).tz_localize(None)
    except:
        pass

    formatStr="%d %b %Y, %H:%M:%S"
    try:
        return pd.to_datetime(d, format=formatStr).tz_localize(None)
    except:
        pass

def convertDates(_startTime, _endTime, meetingData):

    return meetingStartTime, meetingEndTime, meetingData

def meetingIntervals(meetingData, meetingCode, meetingName, meetingStartTime, meetingEndTime, startStep, startRoundDir, midThr, midStep, midRoundDir, endStep, endRoundDir):
    meetingData.loc[:,"startTimeClip"]=meetingData["startTime"].clip(lower=meetingStartTime,upper=meetingEndTime)
    meetingData.loc[:,"endTimeClip"]=meetingData["endTime"].clip(lower=meetingStartTime,upper=meetingEndTime)
    meetingData=meetingData[meetingData["startTimeClip"]!=meetingData["endTimeClip"]]
    try:
        logging.debug("meetingData\n{}".format(meetingData[["Nome partecipante","startTime","endTime"]]))

        participantsTime=meetingData.groupby(["Nome partecipante"])
    except KeyError:
        logging.debug("meetingData\n{}".format(meetingData[["Nome dell'attore","startTime","endTime"]]))

        participantsTime=meetingData.groupby(["Nome dell'attore"])

    results = []

    for name,g in participantsTime:
        participantName = name
        participantRes={
                "Nome": participantName,
                "Assenza intermedia sogliata":0,
                "Assenza intermedia arrotondata":0,
                "Assenza intermedia":0,
                "Presente netto":0,
                "Evento": meetingName
                }
        beginning=True
        logging.debug("participantName {}".format(participantName))
        intervals = []
        for itt,tt in g.sort_values(by="startTime").iterrows():
            if len(intervals) == 0 or tt["startTime"] > intervals[-1]["endTime"]:
                intervals.append(tt[["startTime", "endTime", "startTimeClip","endTimeClip"]])
            else:
                logging.warning("overlapping intervals {} {} and {} {}".format(intervals[-1]["startTime"],intervals[-1]["endTime"],tt["startTime"], tt["endTime"]))
                intervals[-1]["endTimeClip"]=max(intervals[-1]["endTimeClip"],tt["endTimeClip"])
                intervals[-1]["endTime"]=max(intervals[-1]["endTime"],tt["endTime"])

        curTime=meetingStartTime
        for iv in intervals:
            logging.debug("curTime {} startTimeClip {} endTimeClip {}".format(curTime, iv["startTimeClip"],iv["endTimeClip"]))
            diffTime=(iv["startTimeClip"]-curTime).seconds
            logging.debug("diffTime {}".format(diffTime))
            if beginning:
                beginning=False
                participantRes["Orario entrata reale"]=iv["startTime"]
                participantRes["Orario entrata"]=iv["startTimeClip"]
                participantRes["Assenza inizio"]=diffTime
                participantRes["Assenza inizio arrotondata"]=roundTime(diffTime, startStep, startRoundDir)
            else:
                participantRes["Assenza intermedia sogliata"] += diffTime if diffTime >= midThr else 0
                participantRes["Assenza intermedia"] += diffTime
            participantRes["Presente netto"] += (iv["endTimeClip"]-iv["startTimeClip"]).seconds
            curTime=iv["endTimeClip"]
        diffTime=meetingEndTime-curTime
        participantRes["Assenza fine"]=diffTime.seconds
        participantRes["Assenza fine arrotondata"]=roundTime(diffTime.seconds, endStep, endRoundDir)
        participantRes["Orario uscita reale"]=iv["endTime"]
        participantRes["Orario uscita"]=iv["endTimeClip"]
        participantRes["Assenza intermedia arrotondata"] = roundTime(participantRes["Assenza intermedia sogliata"], midStep, midRoundDir)
        participantRes["Presente complessivo arrotondato"]=(meetingEndTime - meetingStartTime).seconds - participantRes["Assenza inizio arrotondata"] - participantRes["Assenza fine arrotondata"] - participantRes["Assenza intermedia arrotondata"]
        participantRes["Codice riunione"]=meetingCode
        participantRes["Orario inizio evento"]=meetingStartTime
        participantRes["Orario fine evento"]=meetingEndTime
        
        if (participantRes["Presente complessivo arrotondato"] > 0):
            results.append(participantRes)

    return results
    

def meetReport (fileName, meetFileData, icsFileData, meetingCode, _startTime, _endTime, startStep, startRoundDir, midThr, midStep, midRoundDir, endStep, endRoundDir):

    if re.match(".*\.csv$", fileName):
        meetData = pd.read_csv(meetFileData)
    elif re.match(".*\.xlsx?$", fileName):
        meetData = pd.read_excel(meetFileData)

    cal = None

    logging.debug("icsFileData {}".format(icsFileData))

    if (icsFileData):
        icsData=icsFileData.read()
        if len(icsData) > 0:
            cal = ics.Calendar(icsData.decode("utf-8"))

    results=[]

    meetData["Data"]=list(map(convertDate,meetData["Data"]))
    meetData["endTime"]=meetData["Data"]

    meetData["startTime"]=list(map(lambda x:x[1]["endTime"]-pd.to_timedelta(x[1]["Durata"],'s'),meetData.iterrows()))

    meetingCodes = set(meetData["Codice riunione"])

    for meetingCode in meetingCodes:
        logging.info("meetingCode {}".format(meetingCode))
        meetingData = meetData[meetData["Codice riunione"]==meetingCode]

        meetingName="N/A"

        if cal is not None:
            for e in cal.events:
                logging.info("event {} meetingCode {}".format(repr(e), getEventMeetingCode(e)))
                if getEventMeetingCode(e) == meetingCode:
                    calStartTime = pd.to_datetime(e.begin.datetime).tz_convert(TZ).tz_localize(None)
                    calEndTime = pd.to_datetime(e.end.datetime).tz_convert(TZ).tz_localize(None)
                    meetingName=e.name
                    meetingStartTime = calStartTime
                    meetingEndTime = calEndTime
                    logging.info("event {} startTime {} endTime {}".format(repr(e),meetingStartTime, meetingEndTime))
                    results += meetingIntervals(meetingData, meetingCode, meetingName, meetingStartTime, meetingEndTime, startStep, startRoundDir, midThr, midStep, midRoundDir, endStep, endRoundDir)

        else:
            meetingStartTime=convertDate(_startTime)
            meetingEndTime=convertDate(_endTime)

            results += meetingIntervals(meetingData, meetingCode, meetingName, meetingStartTime, meetingEndTime, startStep, startRoundDir, midThr, midStep, midRoundDir, endStep, endRoundDir)
        

    output=io.BytesIO()
    dfResults=pd.DataFrame(results)

    for col in ["Assenza inizio arrotondata", "Assenza intermedia arrotondata", "Assenza fine arrotondata", "Presente complessivo arrotondato"]:
        dfResults[col] = list(map(lambda x:formatSeconds(x),dfResults[col]))

    dfResults.to_excel(output)

    return output


if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.DEBUG)
    logging.info("Open {}".format([sys.argv[1]]))
    with open(sys.argv[1]) as f:
        report = meetReport (sys.argv[1], sys.argv[1], open(sys.argv[2],'rb'), sys.argv[3], sys.argv[4], sys.argv[5], int(sys.argv[6]), sys.argv[7], int(sys.argv[8]), int(sys.argv[9]), sys.argv[10], int(sys.argv[11]), sys.argv[12])
        with open("test.xlsx", "wb") as f:
            f.write(report.getbuffer())
