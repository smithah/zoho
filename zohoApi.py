import requests
import xlrd
# Give the location of the config file
configPath="E:/Automation/PerdixAndInfraExitProcess/Data/Config/Config.xlsx"
 
workbook = xlrd.open_workbook(configPath)
worksheet = workbook.sheet_by_name('Settings')

accessToken = worksheet.cell_value(21,1)
print(accessToken)

auth_token="Zoho-oauthtoken "+accessToken #YOUR_AUTH_TOKEN
org_id="60003159359" #YOUR_ORGANISATION_ID

HEADERS={
    "Authorization":auth_token,
    "orgId":org_id,
    "Content-Type": "application/json;charset=UTF-8"    
}


#This api call return opentickets Count of teh agent
def getTicketsCount(agentId):
    request =None
    ticketsCount = ""
    URL = "https://desk.zoho.in/api/v1/agentsTicketsCount"    
    try:
        PARAMS = {
            'agentIds': agentId
            }
        request=requests.get(url = URL,params = PARAMS, headers = HEADERS)
        print (request.content)
        if request.status_code == 200:
            outputData = request.json()
            ticketsCount = outputData['data'][0]['openCount']
            print('ticketsCount: %s' % ticketsCount)            
    except:
        print('Error occured')    
    return ticketsCount

# This api call will deactivate agent 
# API used: Deactivate agent
def deactivateAgent(agentId):
    request =None    
    URL = "https://desk.zoho.in/api/v1/agents/"+agentId+"/deactivate"    
    try:
        request=requests.post(url = URL, headers = HEADERS)
        print (request.content)            
    except:
        print('Error occured')    
    return 1

# This api reassign open tickets to other agent in the same department
def reassignTickets(fromAgentId, toAgentId, dptId):
    request =None
    try:
        URL = "https://desk.zoho.in/api/v1/agents/"+fromAgentId+"/reassignment"
        DATA = {
            "agentReassignment" : [ {
                "departmentId" : dptId,
                "taskNewOwner" : toAgentId,
                "ticketNewOwner" : toAgentId
                } ]
            }
        request=requests.post(url = URL, headers = HEADERS, json = DATA)
        print (request.content)
        return 1
    except:
        print('Error occured')
        return 0

def saveAgentResponse(agentEmail, outputPath):
    request =None
    URL = "https://desk.zoho.in/api/v1/agents"
    try:
        PARAMS = {
            'searchStr': agentEmail
            }
        request=requests.get(url = URL,params = PARAMS, headers = HEADERS)
        print (request.content)
        if request.status_code == 200:
            outputData = request.json()
            with open(outputPath, 'w') as f:              
                f.writelines(request.text)
        return 1  
    except:
        print('Error occured')    
        return 0
   

# This api call returns agentId for input emailId
# API used: Get agent by email ID
def getAgentId(emailId):
    request =None
    agentId = ""
    URL = "https://desk.zoho.in/api/v1/agents/email/"+emailId    
    try:
        request=requests.get(url = URL, headers = HEADERS)
        print (request.content)
        if request.status_code == 200:
            outputData = request.json()
            agentId = outputData['id']
            print('AgentId: %s' % agentId)       
    except:
        print('Error occured')    
    return agentId  
    

def zohoProcess(agentEmail, preDefinedEmail, outputResponsePath):    
    preEmpAgentId = getAgentId(preDefinedEmail)
    request =None
    empAgentId = ""
    URL = "https://desk.zoho.in/api/v1/agents"    
    try:
        PARAMS = {
            'searchStr': agentEmail
            }
        request=requests.get(url = URL,params = PARAMS, headers = HEADERS)
        print (request.content)
        if request.status_code == 200:
            outputData = request.json()
            empAgentId = outputData['data'][0]['id']
            print('AgentId: %s' % empAgentId)
            print('Status : %s' % outputData['data'][0]['status'])
            # ticket Count
            openTicketsCount = getTicketsCount(empAgentId)
            if int(openTicketsCount) > 0 and len(preEmpAgentId) > 0:
                print("Tickets present need to resassign")
                for dpmtId in outputData['data'][0]['associatedDepartmentIds']:
                    print('DepartmentId: %s' % dpmtId)
                    #Reassign open tickets
                    reassignTickets(empAgentId, preEmpAgentId, dpmtId)
                    
            #Deactivate Agent
            deactivateAgent(empAgentId)
            #Save output Response
            saveAgentResponse(agentEmail, outputResponsePath)       
    except:
        print('Error occured')



#zohoProcess("usertest.o@kinaracapital.com","usertest.o1@kinaracapital.com","h")

