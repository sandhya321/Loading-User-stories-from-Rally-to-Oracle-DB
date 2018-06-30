
# coding: utf-8
#!/usr/bin/python
import sys, os
from pyral import Rally, rallyWorkset, RallyRESTAPIError
#from bleach import clean
import cx_Oracle
import pandas as pd
import time,smtplib,base64
from html.parser import HTMLParser

class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()

conn = cx_Oracle.connect(user=user_name,password=pswd,dsn=dsn_full)
curs = conn.cursor()

server = 'server'
user = 'username'
password = 'password'
recipients = 'email_id'
sender = 'email_id of sender'
SUBJECT = "Script Failure Notification"
c=0
interval=5


def user_stories(prj_name):

    rally = Rally('rally_connection', apikey='your_api', workspace='',project=prj_name,server_ping=False)   
    rally.enableLogging("rally.history.showstories")
    response = rally.get('HierarchicalRequirement')
    lastdatarefresh= str(time.strftime("%b %d %Y %H:%M:%S"))
    a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43,a44,a45,a46,a47,a48,a49,a50,a51,a52,a53,a54,a55,a56,a57,a58,a59,a60,a61,a62,a63,a64 = ([] for i in range(64))
    
    # Delete the rows with projectname = prj_name
    state = 'delete from table where PROJECTNAME = :PROJECTNAME'
    curs.execute(state, {'PROJECTNAME' : prj_name})
    
    for story in response:           
                if(story.Name):
                    e1= str(story.Name)
                    e1 = e1.encode('cp850', errors='replace').decode('cp850')
                    
                else:
                    e1=''

                if(story.Workspace):
                    workID = str(story.Workspace.ObjectID)
                    workID = workID.encode('cp850', errors='replace').decode('cp850')
                else:
                    workID = ''

                if(story.Workspace):
                    workn = str(story.Workspace.Name)
                else:
                    workn=''

                if(story.FormattedID):
                    d = str(story.FormattedID)

                else:
                    d=''

                if(story.AcceptedDate):
                    ad = str(story.AcceptedDate)
                else:
                    ad=''

                if(story.Blocked):
                    bl = str(story.Blocked)
                else:
                    bl=''

                if(story.BlockedReason):
                    blr = str(story.BlockedReason)
                else:
                    blr=''

                if(story.Blocker):
                    blockr= str(story.Blocker.Name)
                    blockr = blockr.encode('cp850', errors='replace').decode('cp850')
                else:
                    blockr=''

                if(story.CreationDate):
                    cred= str(story.CreationDate)
                else:
                    cred=''

                if(story.DefectStatus):
                    ds = str(story.DefectStatus)
                else:
                    ds=''

                if(story.Defects):    
                    for de in story.Defects:
                        df = str(de.Name).encode('cp850', errors='replace').decode('cp850')
                else:
                    df = ''

                if(story.Description):
                    j = str(strip_tags(story.Description))
                    j = str((j[:999] + "..") if len(j) > 75 else j)
                    j= j[:3000]
                    j = j.encode('cp850', errors='replace').decode('cp850')
                else:
                    j=''

                if(story.Expedite):
                    expd = str(story.Expedite)
                else:
                    expd=''

                if(story.FlowState):
                    fsn = str(story.FlowState.Name)
                    fsn = fsn.encode('cp850', errors='replace').decode('cp850')
                else:
                    fsn=''

                if(story.FlowStateChangedDate):
                    fsd = str(story.FlowStateChangedDate)      
                else:
                    fsd=''

                if(story.HasParent):
                    hasp = str(story.HasParent)      
                else:
                    hasp=''

                if(story.InProgressDate):
                    inpd = str(story.InProgressDate)      
                else:
                    inpd=''

                if(story.Iteration):
                    it = str(story.Iteration.Name)
                    it = it.encode('cp850', errors='replace').decode('cp850')
                else:
                    it = ''   

                if(story.LastBuild):
                    lbd = str(story.LastBuild)      
                else:
                    lbd =''

                if(story.LastRun):
                    lrun = str(story.LastRun)      
                else:
                    lrun =''

                if(story.LastUpdateDate):
                    m = str(story.LastUpdateDate)

                else:
                    m=''

                if(story.Milestones):
                    for mi in story.Milestones:
                        milest = str(mi.Name)
                        milest = milest.encode('cp850', errors='replace').decode('cp850')
                else:
                    milest =''

                if(story.Notes):
                    notes = str(strip_tags(story.Notes))
                    notes = notes[:1000]
                    notes = notes.encode('cp850', errors='replace').decode('cp850')

                else:
                    notes=""

                if(story.Owner):
                    owner = str(story.Owner.EmailAddress)
                else:
                    owner= ''

                if(story.Package):
                    pack = str(story.Package)      
                else:
                    pack =''   

                if(story.Project.Name):
                    prjn = str(story.Project.Name)
                    prjn = prjn.encode('cp850', errors='replace').decode('cp850')

                else:
                    prjn=''

                if(story.Release):
                    rel = str(story.Release.Name)
                    rel = rel.encode('cp850', errors='replace').decode('cp850')
                else:
                    rel =''
                if(story.ScheduleState):
                    schs = str(story.ScheduleState)      
                else:
                    schs =''

                if(story.ScheduleStatePrefix):
                    schsp = str(story.ScheduleStatePrefix)      
                else:
                    schsp =''

                if(story.Successors):
                    for su in story.Successors:
                        succ = str(su.Name).encode('cp850', errors='replace').decode('cp850')
                    
                else:
                    succ =''
                if(story.Tags):
                    for t in story.Tags:
                        tags = str(t.Name)
                        tags = tags[:2000]
                        tags = tags.encode('cp850', errors='replace').decode('cp850')
                else:
                    tags =''
                if(story.TaskActualTotal):
                    taskat = str(story.TaskActualTotal)      
                else:
                    taskat =''

                if(story.TaskEstimateTotal):
                    tasket = str(story.TaskEstimateTotal)      
                else:
                    tasket =''
                if(story.TaskRemainingTotal):
                    taskrt = str(story.TaskRemainingTotal)      
                else:
                    taskrt =''
                if(story.TaskStatus):
                    taskst = str(story.TaskStatus)      
                else:
                    taskst =''
                if(story.Tasks):
                    for ta in story.Tasks:
                        tas = str(ta.Name)
                        tas = tas[:3000]
                        tas = tas.encode('cp850', errors='replace').decode('cp850')
                else:
                    tas =''
                if(story.TestCaseCount):
                    taskcc = str(story.TestCaseCount)      
                else:
                    taskcc =''
                if(story.TestCaseStatus):
                    taskcs = str(story.TestCaseStatus)      
                else:
                    taskcs =''
                if(story.TestCases):
                    for tcc in story.TestCases:
                        taskc = str(tcc.Name)
                        taskc = taskc[:3000]
                        taskc = taskc.encode('cp850', errors='replace').decode('cp850')
      
                else:
                    taskc =''

                if(story.AcceptanceCriteria):
                    accepc = str(strip_tags(story.AcceptanceCriteria))
                    accepc = accepc[:2000]
                    accepc=accepc.encode('cp850', errors='replace').decode('cp850')
                else:
                    accepc =''

                if(story.AnalysisStatus):
                    analysist = str(story.AnalysisStatus)      
                else:
                    analysist =''

                if(story.AnalysisType):
                    antype = str(story.AnalysisType)      
                else:
                    antype =''

                if(story.BizPriority):
                    bpri = str(story.BizPriority)      
                else:
                    bpri =''
                #if((prj_name == 'NPI Repair Enablement') or (prj_name == 'Service Delivery ART') or(prj_name == 'GSAR PEGA APPS - S2S,DRD, ARCH, FAST') or (prj_name == 'SDR12 PT Object & Static Data Migration') or (prj_name == 'SDR12 PT Foundational Services Team') or (prj_name == 'SDR12 PT Internal & External Services Team') \
                 #    or (prj_name == 'GSLO Finance Team') or (prj_name == 'GSLO Internal & External Services Team')):
                    
                baname =''
                #else:
                 #   if(story.BusinessApplicationName):
                 #       baname = str(story.BusinessApplicationName)  
                  #  
                  #  else:
                  #      baname =''

                if(story.BusinessApplicationNameOld):
                    banameold = str(story.BusinessApplicationNameOld)      
                else:
                    banameold =''

                if(story.BusinessPriority):
                    busp  = str(story.BusinessPriority)      
                else:
                    busp =''
                if(story.BusinessValue):
                    busv  = str(story.BusinessPriority)      
                else:
                    busv =''

                if(story.BusinessValueStatement):
                    bvs = str(story.BusinessValueStatement)      
                else:
                    bvs =''
                if(story.CDRMTest):
                    cdrm = str(story.CDRMTest)      
                else:
                    cdrm =''
                if(story.CrossFunctionalImpact):
                    cfi = str(story.CrossFunctionalImpact)      
                else:
                    cfi =''

                if(story.ERMOReleaseEvent):
                    ermre = str(story.ERMOReleaseEvent)      
                else:
                    ermre =''

                if(story.ERMOReleaseName):
                    ermrn = str(story.ERMOReleaseName)      
                else:
                    ermrn =''
                if(story.EstDevCompleteDate):
                    estdc = str(story.EstDevCompleteDate)      
                else:
                    estdc =''

                if(story.OTBRequestPriority):
                    otbreqpr = str(story.OTBRequestPriority)      
                else:
                    otbreqpr =''

                if(story.ProcessPhase):
                    proph = str(story.ProcessPhase)      
                else:
                    proph =''
                if(story.ProductCategory):
                    procat = str(story.ProductCategory)      
                else:
                    procat =''
                if(story.ProductOwner):
                    proow = str(story.ProductOwner)      
                else:
                    proow =''
                if(story.QAEstimate):
                    qaest = str(story.QAEstimate)      
                else:
                    qaest =''
                if(story.QAStatus):
                    qasta = str(story.QAStatus)      
                else:
                    qasta =''
                if(story.Requester):
                    requ = str(story.Requester)      
                else:
                    requ =''

                if(story.TestCycle):
                    testc = str(story.TestCycle)      
                else:
                    testc =''
                if(story.Track):
                    track = str(story.Track)      
                else:
                    track =''
                if(story.XFlow):
                    xfl = str(story.XFlow)      
                else:
                    xfl =''
                a1.append(e1),a2.append(workID),a3.append(workn),a4.append(d),a5.append(ad),a6.append(bl),a7.append(blr),a8.append(blockr),a9.append(cred),a10.append(ds),a11.append(df),a12.append(j),a13.append(expd),a14.append(fsn),a15.append(fsd),a16.append(hasp),a17.append(inpd),a18.append(it),a19.append(lbd),a20.append(lrun),a21.append(m),a22.append(milest),a23.append(notes),a24.append(owner),a25.append(pack),a26.append(prjn),a27.append(rel),a28.append(schs),a29.append(schsp),a30.append(succ),a31.append(tags),a32.append(taskat),a33.append(tasket),a34.append(taskrt),a35.append(taskst),a36.append(tas),a37.append(taskcc),a38.append(taskcs),a39.append(taskc),a40.append(accepc),a41.append(analysist),a42.append(antype),a43.append(bpri),a44.append(baname),a45.append(banameold),a46.append(busp),a47.append(busv),a48.append(bvs),a49.append(cdrm),a50.append(cfi),a51.append(ermre),a52.append(ermrn),a53.append(estdc),a54.append(otbreqpr),a55.append(proph),a56.append(procat),a57.append(proow),a58.append(qaest),a59.append(qasta),a60.append(requ),a61.append(testc),a62.append(track),a63.append(xfl)
                a64.append(lastdatarefresh)
    
    df2 = pd.DataFrame(list(zip(a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43,a44,a45,a46,a47,a48,a49,a50,a51,a52,a53,a54,a55,a56,a57,a58,a59,a60,a61,a62,a63,a64)))
    
    row_values1= [list(x) for x in df2.values]
    
    
    try:
        curs.executemany ("INSERT INTO rally_User_Stories (User_Story_Name,WorkspaceID,WorkspaceName,FormattedID,AcceptedDate,Blocked,BlockedReason,Blocker,CreationDate,DefectStatus,Defects,Description,Expedite,FlowState,FlowStateChangedDate,HasParent,InProgressDate,Iteration,LastBuild,LastRun,LastUpdateDate,Milestones,Notes,OwnerEmail,PackageN,ProjectName,Release,ScheduleState,ScheduleStatePrefix,Successors,Tags,TaskActualTotal,TaskEstimateTotal,TaskRemainingTotal,TaskStatus,Tasks,TestCaseCount,TestCaseStatus,TestCases,AcceptanceCriteria,AnalysisStatus,AnalysisType,BizPriority,BusinessApplicationName,BusinessApplicationNameOld,BusinessPriority,BusinessValue,BusinessValueStatement,CDRMTest,CrossFunctionalImpact,ERMOReleaseEvent,ERMOReleaseName,EstDevCompleteDate,OTBRequestPriority,ProcessPhase,ProductCategory,ProductOwner,QAEstimate,QAStatus,Requester,TestCycle,Track,XFlow,LastDataRefresh) VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15,:16,:17,:18,:19,:20,:21,:22,:23,:24,:25,:26,:27,:28,:29,:30,:31,:32,:33,:34,:35,:36,:37,:38,:39,:40,:41,:42,:43,:44,:45,:46,:47,:48,:49,:50,:51,:52,:53,:54,:55,:56,:57,:58,:59,:60,:61,:62,:63,:64)",row_values1)
        print('yes')
               
    except Exception as e:
    
        print("err")
        conn.rollback()

        global c
        global SUBJECT
        c=c+1
        session = smtplib.SMTP()
        session.connect(server,25)
        session.ehlo()
        session.starttls()
        session.ehlo()
        # if your SMTP server doesn't need authentications,
        # you don't need the following line:
        #session.login(base64.b64decode(user),base64.b64decode(password))
        TEXT = "Hello! \n\n The incremental User Stories script with error handling ("+str(c)+")  has failed for the interval of "+str(interval)+" minutes with the following exception/error:\n\n\nERROR: "+ str(e) +". Error in ProjectName :  "+ str(prj_name) + "\n\nThis may impact the Loading of the data from Rally into Oracle.\n\n\n Regards,\n\n"
        message = 'Subject: {}\n\n{}'.format(SUBJECT, TEXT)
        session.sendmail(sender, recipients, message)

    else:
        conn.commit()

def main():
   
    
    user_stories('prj_name')

if __name__ == "__main__":
    main()

curs.close()
conn.close()

