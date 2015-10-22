// Task Scheluder 2.0.cpp : Defines the entry point for the console application.
//
// Titulo: prueba del API 2.0 de Task Scheluder 
// Autor : Sheng
// Fecha : 04/03/2013
// Descripcion: 
/*    To display task names and state for all the tasks in a task folder
		1.- Initialize COM and set general COM security.
		2.- Create the ITaskService object.
			This object allows you to connect to the Task Scheduler service and access a specific task folder.
		3.- Get a task folder that holds the tasks you want information about.
			Use the ITaskService::GetFolder method to get the folder.
		4.- Get the collection of tasks from the folder.
			Use the ITaskFolder::GetTasks method to get the collection of tasks (IRegisteredTaskCollection).
		5.- Get the number of tasks in the collection, and enumerate through each task in the collection.
			Use the Item Property of IRegisteredTaskCollection to get an IRegisteredTask instance. 
			Each instance will contain a task in the collection. You can then display the information 
			(property values) from each registered task.
  */



#define _WIN32_DCOM

#include "stdafx.h"
#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <string.h>

#include <taskschd.h> // Include the task header file.


#pragma comment(lib,"taskschd.lib")
#pragma comment(lib,"comsupp.lib")

using namespace std; 

string getTaskInstancesPolicyString(TASK_INSTANCES_POLICY n);
string getCompatibilityString(TASK_COMPATIBILITY  n);
string getBoolString (VARIANT_BOOL cond);
string getRunLevelString (int n); 
string getLogonTypeString (int n);
string getState(int n);
string getTypeTriggerString (TASK_TRIGGER_TYPE2  n); 

void getITaskSettings(ITaskSettings *ppSettings);
void getIPrincipal(IPrincipal  *ppPrincipal); 
void getIRegistrationInfo(IRegistrationInfo   *ppRegistrationInfo); 
void getITriggerCollection(ITriggerCollection *pTriggers);


// Identificacion de conexion, si es localhost no hace falta ningun dato.
const char TASK_SERVER_NAME [50] = "192.168.0.99"; 
const char TASK_USER[50]         = "Administrator"; 
const char TASK_DOMAIN[50]       = "EMEAW2K8DMN.COM"; 
const char TASK_PASSWORD[50]     = "Tango1234"; 




int _tmain(int argc, _TCHAR* argv[])
{
	// Inicializamos el COM 
	HRESULT hr = CoInitializeEx (NULL, COINIT_MULTITHREADED);

	if (FAILED(hr)) 
	{
		wcout << "\n CoInitializeEx failed: " << hr << endl;
		return 1; 
	}
	//  ------------------------------------------------------
	// Set general COM Security levels.
	hr = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		0,
		NULL);

	if( FAILED(hr) )
	{
		wcout << "\nCoInitializeSecurity failed: " << hr << endl;
		CoUninitialize();   
		return 1;
	}
	//  ------------------------------------------------------
	//  Create an instance of the Task Service. 
	//  La única dif. con la anterior API, se instancia una clase ITaskService en vez de ITaskScheduler 
	ITaskService *pService = NULL;
	hr = CoCreateInstance( CLSID_TaskScheduler,
						   NULL,
						   CLSCTX_INPROC_SERVER,
						   IID_ITaskService,
						   (void**)&pService );  

	if (FAILED(hr))
	{
		  wcout << "Failed to CoCreate an instance of the TaskService class:" <<  hr << endl;
		  CoUninitialize();
		  return 1;
	}
	//  ------------------------------------------------------
	/*****************************************************************************************************
	 * 
	 * The ITaskService::Connect method should be called before calling any of the other ITaskService methods.
	 *
     * If you are to connecting to a remote Windows Vista computer from a Windows Vista, you need to allow the 
	 * Remote Scheduled Tasks Management firewall exception on the remote computer. To allow this exception 
	 * click Start, Control Panel, Security, Allow a program through Windows Firewall, and then select the 
	 * Remote Scheduled Tasks Management check box. Then click the Ok button in the Windows Firewall Settings 
	 * dialog box.
	 *
     * If you are connecting to a remote Windows XP or Windows Server 2003 computer from a Windows Vista computer, 
	 * you need to allow the File and Printer Sharing firewall exception on the remote computer. To allow this 
	 * exception click Start, Control Panel, double-click Windows Firewall, select the Exceptions tab, and then 
	 * select the File and Printer Sharing firewall exception. Then click the OK button in the Windows Firewall 
	 * dialog box. The Remote Registry service must also be running on the remote computer.
	 *
	 ********************************************************************************************************/
	//  Connect to the task service.
	hr = pService->Connect(_variant_t(TASK_SERVER_NAME),    //ServerName
						   _variant_t(TASK_USER),           //User
						   _variant_t(TASK_DOMAIN),         //domain 
						   _variant_t(TASK_PASSWORD));      //password
	if( FAILED(hr) )
	{
		wcout << "ITaskService::Connect failed: " << hr << endl; 
		pService->Release();
		CoUninitialize();
		return 1;
	}
	//  ------------------------------------------------------
	//  Get the pointer to the root task folder.
	ITaskFolder *pRootFolder = NULL;
	hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
	
	pService->Release();
	if( FAILED(hr) )
	{
		wcout << "Cannot get Root Folder pointer: " <<  hr << endl;
		CoUninitialize();
		return 1;
	}
	//  ------------------------------------------------------
	//  Get the registered tasks in the folder.
	IRegisteredTaskCollection* pTaskCollection = NULL;
	hr = pRootFolder->GetTasks( NULL, &pTaskCollection );

	pRootFolder->Release();
	if( FAILED(hr) )
	{
		wcout << "Cannot get the registered tasks.: " << hr << endl;
		CoUninitialize();
		return 1;
	}

	LONG numTasks = 0;
	hr = pTaskCollection->get_Count(&numTasks);

	if( numTasks == 0 )
	 {
		wcout << "No Tasks are currently running" << endl;
		pTaskCollection->Release();
		CoUninitialize();
		return 1;
	 }

	 wcout << "Number of Tasks: " << numTasks << endl;

	TASK_STATE taskState;
	
	for(LONG i=0; i < numTasks; i++)
	{
		IRegisteredTask* pRegisteredTask = NULL;
		hr = pTaskCollection->get_Item( _variant_t(i+1), &pRegisteredTask );
		
		if( SUCCEEDED(hr) )
		{
			// TASK NAME
			BSTR taskName = NULL;
			hr = pRegisteredTask->get_Name(&taskName);
			if( SUCCEEDED(hr) )
			{
				wcout << "Task Name: " << (taskName) << endl; 
				SysFreeString(taskName);
				
				// TASK STATES 
				hr = pRegisteredTask->get_State(&taskState);
				if (SUCCEEDED (hr) )
				{			
					wcout << "\tState: " << (getState(taskState)).c_str() << endl;
				}
				else 
				{
					wcout << "\tState: " << endl;
				}
				// Enabled 
				VARIANT_BOOL enabled; 
				hr = pRegisteredTask->get_Enabled(&enabled);
				if (SUCCEEDED(hr)) 
				{
					wcout << "\tEnabled: " << getBoolString(enabled).c_str() << endl; 
				}
				// LastRunTime
				DATE LastRunTime; 
				hr = pRegisteredTask->get_LastRunTime(&LastRunTime);
				if (SUCCEEDED(hr))
				{
					SYSTEMTIME systemTime; 

					VariantTimeToSystemTime(LastRunTime,&systemTime); 
					wcout << "\tLastRunTime : " << systemTime.wDay << "/" 
												<< systemTime.wMonth << "/" 
												<< systemTime.wYear << "  " 
												<< systemTime.wHour << ":" 
												<< systemTime.wMinute << ":" 
												<< systemTime.wSecond << endl; 
				}
				// LastTaskResult
				HRESULT LastTaskResult; 
				hr = pRegisteredTask->get_LastTaskResult(&LastTaskResult);
				if (SUCCEEDED(hr))
				{
					wcout << "\tLastTaskResult: " << LastTaskResult << endl; 
				}
				// NextRunTime
				DATE NextRunTime; 
				hr = pRegisteredTask->get_NextRunTime(&NextRunTime);
				if (SUCCEEDED(hr))
				{
					SYSTEMTIME systemTime; 

					VariantTimeToSystemTime(NextRunTime,&systemTime); 
					wcout << "\tNextRunTime : " << systemTime.wDay << "/" 
												<< systemTime.wMonth << "/" 
												<< systemTime.wYear << "  " 
												<< systemTime.wHour << ":" 
												<< systemTime.wMinute << ":" 
												<< systemTime.wSecond << endl; 
				}
				// NumberOfMissedRuns
				LONG NumberOfMissedRuns; 
				hr = pRegisteredTask->get_NumberOfMissedRuns(&NumberOfMissedRuns);
				if (SUCCEEDED(hr))
				{
					wcout << "\tNumberOfMissedRuns : " << NumberOfMissedRuns << endl; 
				}
				// Path
				BSTR Path = NULL;
				hr = pRegisteredTask->get_Path(&Path);
				if ((SUCCEEDED(hr)) && (Path != NULL))
				{
					wcout << "\tPath : " << Path << endl; 
					SysFreeString(Path);
				}

				//--------------------- DEFINITION --------------------------------------
				ITaskDefinition *ppDefinition = NULL; 
				hr = pRegisteredTask->get_Definition(&ppDefinition);

				if (SUCCEEDED (hr))
				{
					//---------------- DATA              ---------------------------------
					BSTR pData = NULL; 
					hr = ppDefinition->get_Data(&pData); 
					if ((SUCCEEDED (hr)) && (pData != NULL))
					{
						wcout << "\tData: " << pData << endl; 
						SysFreeString(pData);
					}
					else
					{
						wcout << "\tData: " << endl; 
					}
					//---------------- REGISTRATION INFO ---------------------------------
					IRegistrationInfo *ppRegistrationInfo = NULL;
					hr = ppDefinition->get_RegistrationInfo(&ppRegistrationInfo);					
					if (SUCCEEDED (hr))
					{
						getIRegistrationInfo(ppRegistrationInfo); 
					}
					else
					{
						wcout << "\tCannot get the RegistrationInfo: " << hr << endl; 
					}
					ppRegistrationInfo->Release();
					//------------------- PRINCIPAL INFO ---------------------------------
					IPrincipal  *ppPrincipal  = NULL;
					hr = ppDefinition->get_Principal(&ppPrincipal);					
					if (SUCCEEDED (hr))
					{
						getIPrincipal(ppPrincipal);
					}
					else
					{
						wcout << "\tCannot get the IPrincipal: " << hr << endl; 
					}
					ppPrincipal->Release();
					//------------------- SETTING INFO ---------------------------------
					ITaskSettings  *ppSettings  = NULL;
					hr = ppDefinition->get_Settings(&ppSettings);					
					if (SUCCEEDED (hr))
					{
						getITaskSettings(ppSettings);
					}
					else
					{
						wcout << "\tCannot get the ITaskSettings: " << hr << endl; 
					}
					ppSettings->Release();
					//------------------- Triggers INFO ---------------------------------
					ITriggerCollection   *pTriggers  = NULL;
					hr = ppDefinition->get_Triggers(&pTriggers);					
					if (SUCCEEDED (hr))
					{
						getITriggerCollection(pTriggers);
					}
					else
					{
						wcout << "\tCannot get the ITriggerCollection: " << hr << endl; 
					}
					pTriggers->Release();
				}
				else
				{
					wcout << "\tCannot get the TaskDefinition: " << hr << endl; 
				}
				ppDefinition->Release();
			}
			else
			{
				wcout << "Cannot get the registered task name: " << hr << endl;
			}
			pRegisteredTask->Release();
		}
		else
		{
			wcout << "Cannot get the registered task item at index= " << i+1 << ": " <<  hr << endl;
		}
	}

	pTaskCollection->Release();
	CoUninitialize();
	return 0;
}
/*****************************************************************************************************
 * RegistrationInfo Interface
 *
 *   Provides the administrative information that can be used to describe the task. 
 * This information includes details such as a description of the task, the author of the task, the 
 * date the task is registered, and the security descriptor of the task.
 *
 *   Property: 
 *     - Author: the author of the task.
 *     - Date : the date and time when the task is registered.
 *     - Description: the description of the task.
 *     - Documentation: any additional documentation for the task.
 *     - SecurityDescriptor: the security descriptor of the task.
 *     - Source:  where the task originated from. For example, a task may originate from a component, service, application, or user.
 *     - URI: the URI of the task.
 *     - Version: the version number of the task.
 *     - XmlText: an XML-formatted version of the registration information for the task.
 *
 *  Link: 
 *     - http://msdn.microsoft.com/es-es/library/windows/desktop/aa380773(v=vs.85).aspx
 *****************************************************************************************************/ 
void getIRegistrationInfo(IRegistrationInfo   *ppRegistrationInfo)
{
	HRESULT hr; 

	wcout << "\t IRegistrationInfo " << endl; 
	wcout << "\t ----------------- " << endl; 

	// Creator: Gets the author of the task
	BSTR pAuthor = NULL; 
	hr = ppRegistrationInfo->get_Author (&pAuthor);
	if ((SUCCEEDED (hr)) && (pAuthor != NULL))
	{
		wcout << "\tCreator: " <<  (pAuthor) << endl; 
		SysFreeString(pAuthor);
	}
	else
	{
		wcout << "\tCreator: " << endl; 
	}
	// Data created: Gets the date and time when the task is registered.
	BSTR pDate = NULL; 
	hr = ppRegistrationInfo->get_Date (&pDate);
	if ((SUCCEEDED (hr)) && (pDate != NULL))
	{
		wcout << "\tDate created: " <<  (pDate) << endl; 
		SysFreeString(pDate);
	}
	else
	{
		wcout << "\tDate created: " << endl; 
	}
	// Description: Gets he description of the task.
	BSTR description = NULL; 
	hr = ppRegistrationInfo->get_Description (&description);
	if ((SUCCEEDED (hr)) && (description != NULL))
	{
		wcout << "\tDescription: " <<  (description) << endl; 
		SysFreeString(description);
	}
	else
	{
		wcout << "\tDescription: " << endl; 
	}
	// Source: Gets the URI of the task.
	BSTR pUri = NULL; 
	hr = ppRegistrationInfo->get_Source (&pUri);
	if ((SUCCEEDED (hr)) && (NULL != pUri))
	{
		wcout << "\tSource or URI: " <<  (pUri) << endl; 
		SysFreeString(pUri);
	}
	else
	{
		wcout << "\tpUri: " << endl; 
	}						
	// XMLText: Gets the URI of the task.
	BSTR XMLText = NULL; 
	hr = ppRegistrationInfo->get_XmlText (&XMLText);
	if ((SUCCEEDED (hr)) && (XMLText != NULL))
	{
		wcout << "\tXMLText: " <<  (XMLText) << endl; 
		SysFreeString(XMLText);
	}
	else
	{
		wcout << "\tXMLText: " << endl; 
	}
}
/*****************************************************************************************************
 * IPrincipal Interface
 *
 *   Provides the security credentials for a principal. These security credentials define the 
 * security context for the tasks that are associated with the principal.
 *
 *   Property: 
 *     - DisplayName
 *     - GroupId
 *     - Id
 *     - LogonType
 *     - RunLevel
 *     - UserId
 *
 *  Link: 
 *     - http://msdn.microsoft.com/es-es/library/windows/desktop/aa380742(v=vs.85).aspx
 *****************************************************************************************************/ 
void getIPrincipal(IPrincipal  *ppPrincipal)
{
	HRESULT hr; 

	wcout << "\tIPrincipal " << endl; 
	wcout << "\t---------- " << endl; 

	// DisplayName: 
	BSTR pName = NULL; 
	hr = ppPrincipal->get_DisplayName (&pName);
	if ((SUCCEEDED (hr)) && (pName != NULL))
	{
		wcout << "\tDisplayName: " <<  (pName) << endl; 
		SysFreeString(pName);
	}
	else
	{
		wcout << "\tDisplayName: " << endl; 
	}
	// Goup: user group that is required to run the tasks that are associated with the principal.
	BSTR pGroup = NULL; 
	hr = ppPrincipal->get_GroupId (&pGroup);
	if ((SUCCEEDED (hr)) && (pGroup != NULL))
	{
		wcout << "\tGroup: " <<  (pGroup) << endl; 
		SysFreeString(pGroup);
	}
	else
	{
		wcout << "\tGroup: " << endl; 
	}
	// ID: the identifier of the principal
	BSTR pId = NULL; 
	hr = ppPrincipal->get_Id (&pId);
	if ((SUCCEEDED (hr)) && (pId != NULL))
	{
		wcout << "\tId: " <<  (pId) << endl; 
		SysFreeString(pId);
	}
	else
	{
		wcout << "\tId: " << endl; 
	}
	// LogonType: the security logon method that is required to run the tasks that are associated with the principal.
	TASK_LOGON_TYPE pLogon; 
	hr = ppPrincipal->get_LogonType (&pLogon);
	if ((SUCCEEDED (hr)) && (pLogon != NULL))
	{
		wcout << "\tTask Logon Type: " <<  getLogonTypeString(pLogon).c_str() << endl; 
	}
	else
	{
		wcout << "\tTask Logon Type: " << endl; 
	}
	// RunLevel: the identifier that is used to specify the privilege level that is required to run the tasks that are associated with the principal.
	TASK_RUNLEVEL_TYPE  pRunLevel; 
	hr = ppPrincipal->get_RunLevel (&pRunLevel);
	if ((SUCCEEDED (hr)) && (pRunLevel != NULL))
	{
		wcout << "\tTask RunLevel: " <<  getRunLevelString(pRunLevel).c_str() << endl; 
	}
	else
	{
		wcout << "\tTask RunLevel: " << endl; 
	}
	// UserID: the user identifier that is required to run the tasks that are associated with the principal.
	BSTR   pUserId; 
	hr = ppPrincipal->get_UserId (&pUserId);
	if ((SUCCEEDED (hr)) && (pUserId != NULL))
	{
		wcout << "\tUser Id: " <<  pUserId << endl;
		SysFreeString(pUserId);
	}
	else
	{
		wcout << "\tUser Id: " << endl; 
	}

}
/*****************************************************************************************************
 * RegistrationInfo Interface
 *
 *   Provides the settings that the Task Scheduler service uses to perform the task.
 *
 *   Property: 
 *     - AllowDemandStart:  a Boolean value that indicates that the task can be started by using either the Run command or the Context menu.
 *     - AllowHardTerminate:  a Boolean value that indicates that the task may be terminated by using TerminateProcess.
 *     - Compatibility:  an integer value that indicates which version of Task Scheduler a task is compatible with.
 *     - DeleteExpiredTaskAfter:  the amount of time that the Task Scheduler will wait before deleting the task after it expires.
 *     - DisallowStartIfOnBatteries:  a Boolean value that indicates that the task will not be started if the computer is running on battery power.
 *     - Enabled:  a Boolean value that indicates that the task is enabled. The task can be performed only when this setting is true.
 *     - ExecutionTimeLimit:  the amount of time that is allowed to complete the task.
 *     - Hidden:  a Boolean value that indicates that the task will not be visible in the UI by default.
 *     - IdleSettings:  the information that specifies how the Task Scheduler performs tasks when the computer is in an idle state.
 *     - MultipleInstances:  the policy that defines how the Task Scheduler handles multiple instances of the task.
 *     - NetworkSettings:  the network settings object that contains a network profile identifier and name. If the RunOnlyIfNetworkAvailable property of ITaskSettings is true and a network propfile is specified in the NetworkSettings property, then the task will run only if the specified network profile is available.
 *     - Priority:  the priority level of the task.
 *     - RestartCount:  the number of times that the Task Scheduler will attempt to restart the task.
 *     - RestartInterval:  a value that specifies how long the Task Scheduler will attempt to restart the task.
 *     - RunOnlyIfNetworkAvailable:  a Boolean value that indicates that the Task Scheduler will run the task only when a network is available.
 *     - StartWhenAvailable:  a Boolean value that indicates that the Task Scheduler can start the task at any time after its scheduled time has passed.
 *     - StopIfGoingOnBatteries:  a Boolean value that indicates that the task will be stopped if the computer switches to battery power.
 *     - WakeToRun:  a Boolean value that indicates that the Task Scheduler will wake the computer before it runs the task.
 *     - XmlText:  an XML-formatted definition of the task settings.
 *
 *  Link: 
 *     - http://msdn.microsoft.com/en-us/library/windows/desktop/aa381843(v=vs.85).aspx
 *****************************************************************************************************/ 
void getITaskSettings(ITaskSettings *ppSettings)
{
	HRESULT hr; 

	wcout << "\t ITaskSettings " << endl; 
	wcout << "\t ------------- " << endl; 

	// AllowDemandStart 
	VARIANT_BOOL  pAllowDemandStart; 
	hr = ppSettings->get_AllowDemandStart (&pAllowDemandStart);
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t AllowDemandStart: " <<  getBoolString(pAllowDemandStart).c_str()  << endl; 		
	}
	else
	{
		wcout << "\t AllowDemandStart: " << endl; 
	}
	// AllowHardTerminate 
	VARIANT_BOOL  pAllowHardTerminate; 
	hr = ppSettings->get_AllowHardTerminate (&pAllowHardTerminate);
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t AllowHardTerminate: " <<  getBoolString(pAllowHardTerminate).c_str()  << endl; 		
	}
	else
	{
		wcout << "\t AllowHardTerminate: " << endl; 
	}
	// Compatibility 
	TASK_COMPATIBILITY pVer; 
	hr = ppSettings->get_Compatibility (&pVer);
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t Compatibility : " <<  getCompatibilityString(pVer).c_str() << endl; 		
	}
	else
	{
		wcout << "\t Compatibility : " << endl; 
	}
	// DeleteExpiredTaskAfter
	BSTR  pExpirationDelay; 
	hr = ppSettings->get_DeleteExpiredTaskAfter (&pExpirationDelay);
	if ((SUCCEEDED (hr)) && (pExpirationDelay != NULL))
	{
		wcout << "\t DeleteExpiredTaskAfter : " <<  pExpirationDelay << endl; 		
		SysFreeString(pExpirationDelay);
	}
	else
	{
		wcout << "\t DeleteExpiredTaskAfter : " << endl; 
	}
	// DisallowStartIfOnBatteries 
	VARIANT_BOOL  DisallowStartIfOnBatteries; 
	hr = ppSettings->get_DisallowStartIfOnBatteries (&DisallowStartIfOnBatteries);
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t DisallowStartIfOnBatteries: " <<  getBoolString(DisallowStartIfOnBatteries).c_str() << endl; 		
	}
	else
	{
		wcout << "\t DisallowStartIfOnBatteries: " << endl; 
	}
	// Enabled
	VARIANT_BOOL  pEnabled; 
	hr = ppSettings->get_DisallowStartIfOnBatteries (&pEnabled);
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t Enabled: " <<  getBoolString(pEnabled).c_str() << endl; 		
	}
	else
	{
		wcout << "\t Enabled: " << endl; 
	}
	//ExecutionTimeLimit
	BSTR  pExecutionTimeLimit; 
	hr = ppSettings->get_ExecutionTimeLimit (&pExecutionTimeLimit);
	if ((SUCCEEDED (hr)) && (pExecutionTimeLimit != NULL))
	{
		wcout << "\t ExecutionTimeLimit  : " <<  pExecutionTimeLimit << endl; 		
		SysFreeString(pExecutionTimeLimit);
	}
	else
	{
		wcout << "\t ExecutionTimeLimit  : " << endl; 
	}
	// Hidden 
	VARIANT_BOOL  Hidden ; 
	hr = ppSettings->get_Hidden (&Hidden );
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t Hidden : " <<  getBoolString(Hidden).c_str() << endl; 		
	}
	else
	{
		wcout << "\t Hidden : " << endl; 
	}
	// IIdleSettings 
	IIdleSettings *ppIdleSettings = NULL; 
	hr = ppSettings->get_IdleSettings (&ppIdleSettings); 
	if (SUCCEEDED(hr))
	{
		wcout << "\t\t IdleSettings: " << endl; 
		wcout << "\t\t ------------  " << endl; 
		// IdleDuration 
		BSTR  pDelay; 
		hr = ppIdleSettings->get_IdleDuration (&pDelay);
		if ((SUCCEEDED (hr)) && (pDelay != NULL))
		{
			wcout << "\t\t IdleDuration  : " <<  pDelay << endl; 		
			SysFreeString(pDelay);
		}
		else
		{
			wcout << "\t\t IdleDuration  : " << endl; 
		}
		// RestartOnIdle  
		VARIANT_BOOL  restart ; 
		hr = ppIdleSettings->get_RestartOnIdle (&restart );
		if (SUCCEEDED (hr)) 
		{
			wcout << "\t\t RestartOnIdle  : " <<  getBoolString(restart).c_str() << endl; 		
		}
		else
		{
			wcout << "\t\t RestartOnIdle  : " << endl; 
		}
		// StopOnIdleEnd   
		VARIANT_BOOL  stop ; 
		hr = ppIdleSettings->get_StopOnIdleEnd (&stop );
		if (SUCCEEDED (hr)) 
		{
			wcout << "\t\t StopOnIdleEnd   : " <<  getBoolString(stop).c_str() << endl; 		
		}
		else
		{
			wcout << "\t\t StopOnIdleEnd   : " << endl; 
		}
		// WaitTimeout 
		BSTR  timeout; 
		hr = ppIdleSettings->get_WaitTimeout (&timeout);
		if ((SUCCEEDED (hr)) && (timeout != NULL))
		{
			wcout << "\t\t WaitTimeout   : " <<  timeout << endl; 		
			SysFreeString(timeout);
		}
		else
		{
			wcout << "\t\t WaitTimeout   : " << endl; 
		}
	}
	else
	{
	}
	ppIdleSettings->Release(); 
	// MultipleInstances
	TASK_INSTANCES_POLICY   policy ; 
	hr = ppSettings->get_MultipleInstances (&policy);
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t MultipleInstances : " <<  getTaskInstancesPolicyString(policy).c_str() << endl; 		
	}
	else
	{
		wcout << "\t MultipleInstances : " << endl; 
	}
	// NetworkSettings , No vamos a mostrarlo.

	// Priority, default 7 = NORMAL.
	int priority; 
	hr = ppSettings->get_Priority(&priority); 
	if (SUCCEEDED (hr))
	{
		wcout << "\t Priority: " << priority << endl; 
	}
	else
	{
		wcout << "\t Priority: " << endl; 
	}
	// RestartCount 
	int restartCount; 
	hr = ppSettings->get_RestartCount(&restartCount); 
	if (SUCCEEDED (hr))
	{
		wcout << "\t RestartCount: " << restartCount << endl; 
	}
	else
	{
		wcout << "\t RestartCount: " << endl; 
	}
	//RestartInterval 
	BSTR  restartInterval; 
	hr = ppSettings->get_RestartInterval (&restartInterval);
	if ((SUCCEEDED (hr)) && (restartInterval != NULL))
	{
		wcout << "\t RestartInterval  : " <<  restartInterval << endl; 		
		SysFreeString(restartInterval);
	}
	else
	{
		wcout << "\t RestartInterval  : " << endl; 
	}
	// RunOnlyIfNetworkAvailable  
	VARIANT_BOOL runOnlyIfNetworkAvailable ; 
	hr = ppSettings->get_RunOnlyIfNetworkAvailable (&runOnlyIfNetworkAvailable );
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t RunOnlyIfNetworkAvailable : " <<  getBoolString(runOnlyIfNetworkAvailable).c_str() << endl; 		
	}
	else
	{
		wcout << "\t RunOnlyIfNetworkAvailable : " << endl; 
	}
	// StartWhenAvailable  
	VARIANT_BOOL startWhenAvailable ; 
	hr = ppSettings->get_StartWhenAvailable (&startWhenAvailable );
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t StartWhenAvailable : " <<  getBoolString(startWhenAvailable).c_str() << endl; 		
	}
	else
	{
		wcout << "\t StartWhenAvailable : " << endl; 
	}
	// StopIfGoingOnBatteries 
	VARIANT_BOOL stopIfOnBatteries ; 
	hr = ppSettings->get_StopIfGoingOnBatteries (&stopIfOnBatteries );
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t StopIfGoingOnBatteries : " <<  getBoolString(stopIfOnBatteries).c_str() << endl; 		
	}
	else
	{
		wcout << "\t StopIfGoingOnBatteries : " << endl; 
	}
	// WakeToRun 
	VARIANT_BOOL wake ; 
	hr = ppSettings->get_WakeToRun (&wake );
	if (SUCCEEDED (hr)) 
	{
		wcout << "\t WakeToRun : " <<  getBoolString(wake).c_str() << endl; 		
	}
	else
	{
		wcout << "\t WakeToRun : " << endl; 
	}
}
/*****************************************************************************************************
 * RegistrationInfo Interface
 *
 *   Provides the methods that are used to add to, remove from, and get the triggers of a task.
 *
 *   Property: 
 *     - _NewEnum: the collection enumerator for this collection.
 *     - Count: the number of triggers in the collection.
 *     - Item: the specified trigger from the collection.
 *
 *  Link: 
 *     - http://msdn.microsoft.com/en-us/library/windows/desktop/aa381843(v=vs.85).aspx
 *****************************************************************************************************/ 
void getITriggerCollection(ITriggerCollection *pTriggers)
{
	HRESULT hr; 

	wcout << "\t ITriggerCollection " << endl; 
	wcout << "\t ------------------ " << endl; 

	//Num de Triggers
	long count; 
	hr = pTriggers->get_Count (&count); 
	if (SUCCEEDED(hr)) 
	{
		wcout << "\t Num of Triggers: " << count << endl; 
	}

	ITrigger *pTrigger = NULL; 
	for (long i = 0; i < count; i++) 
	{
		hr = pTriggers->get_Item (i + 1, &pTrigger); 
		if ((SUCCEEDED (hr)) && (pTrigger))
		{
			// Enabled 
			VARIANT_BOOL  enabled; 
			hr = pTrigger->get_Enabled (&enabled);
			if (SUCCEEDED (hr)) 
			{
				wcout << "\t\t Enabled: " <<  getBoolString(enabled).c_str()  << endl; 		
			}
			else
			{
				wcout << "\t\t Enabled: " << endl; 
			}
			// EndBoundary 
			BSTR  end; 
			hr = pTrigger->get_EndBoundary (&end);
			if ((SUCCEEDED (hr)) && (end != NULL))
			{
				wcout << "\t\t EndBoundary  : " <<  end << endl; 		
				SysFreeString(end);
			}
			else
			{
				wcout << "\t\t EndBoundary  : " << endl; 
			}
			// ExecutionTimeLimit  
			BSTR  timelimit; 
			hr = pTrigger->get_ExecutionTimeLimit (&timelimit);
			if ((SUCCEEDED (hr)) && (timelimit != NULL))
			{
				wcout << "\t\t ExecutionTimeLimit   : " <<  timelimit << endl; 		
				SysFreeString(timelimit);
			}
			else
			{
				wcout << "\t\t ExecutionTimeLimit   : " << endl; 
			}
			// Id trigger  
			BSTR  id; 
			hr = pTrigger->get_Id (&id);
			if ((SUCCEEDED (hr)) && (id != NULL))
			{
				wcout << "\t\t Id trigger   : " <<  id << endl; 		
				SysFreeString(id);
			}
			else
			{
				wcout << "\t\t Id trigger   : " << endl; 
			}
			// StartBoundary   
			BSTR  start; 
			hr = pTrigger->get_StartBoundary (&start);
			if ((SUCCEEDED (hr)) && (start != NULL))
			{
				wcout << "\t\t StartBoundary : " <<  start << endl; 		
				SysFreeString(start);
			}
			else
			{
				wcout << "\t\t StartBoundary : " << endl; 
			}
			// Type   
			TASK_TRIGGER_TYPE2   type; 
			hr = pTrigger->get_Type (&type);
			if (SUCCEEDED (hr))
			{
				wcout << "\t\t Type : " <<  getTypeTriggerString(type).c_str() << endl; 		
			}
			else
			{
				wcout << "\t\t Type : " << endl; 
			}
			// Repetición
			IRepetitionPattern *pRepeat; 
			hr = pTrigger->get_Repetition (&pRepeat); 
			if (SUCCEEDED (hr)) 
			{
				BSTR duration; 
				hr = pRepeat->get_Duration(&duration); 
				if ((SUCCEEDED (hr)) && (duration != NULL))
				{
					wcout << "\t\t Duration Repetition : " << duration << endl; 
					SysFreeString(duration);
				}
				BSTR interval; 
				hr = pRepeat->get_Interval(&interval);
				if ((SUCCEEDED(hr)) && (interval != NULL))
				{
					wcout << "\t\t Interval Repetition : " << interval << endl; 
					SysFreeString(interval); 
				}
				VARIANT_BOOL  stop; 
				hr = pRepeat->get_StopAtDurationEnd (&stop);
				if (SUCCEEDED (hr)) 
				{
					wcout << "\t\t StopAtDurationEnd : " <<  getBoolString(stop).c_str()  << endl; 		
				}
				else
				{
					wcout << "\t\t StopAtDurationEnd : " << endl; 
				}
			}
			pRepeat->Release(); 
		}

	}
	if (pTrigger!= NULL) pTrigger->Release();     // Liberamos si estaba creado correctamente.
}
string getTaskInstancesPolicyString(TASK_INSTANCES_POLICY n)
{
	switch(n)
	{
		case TASK_INSTANCES_PARALLEL:   return "Starts a new instance while an existing instance of the task is running."; 
		case TASK_INSTANCES_QUEUE:      return "Starts a new instance of the task after all other instances of the task are complete."; 
		case TASK_INSTANCES_IGNORE_NEW: return  "Does not start a new instance if an existing instance of the task is running."; 
		case TASK_INSTANCES_STOP_EXISTING: return "Stops an existing instance of the task before it starts new instance.";
		default: return ""; 
	}
}
string getCompatibilityString(TASK_COMPATIBILITY  n) 
{
	switch(n)
	{
		case TASK_COMPATIBILITY_AT: return "The task is compatible with the AT command.";
		case TASK_COMPATIBILITY_V1: return "The task is compatible with Task Scheduler 1.0."; 
		case TASK_COMPATIBILITY_V2: return "The task is compatible with Task Scheduler 2.0."; 
	}
	return ""; 
}
string getBoolString (VARIANT_BOOL cond)
{
	if (cond == VARIANT_TRUE) 
	{
		return "TRUE";
	}
	else // cond == VARIANT_FALSE
	{
		return "FALSE"; 
	}
	return "FALSE"; 
}
string getRunLevelString (int n) 
{
	switch (n)
	{
		case 0: return "TASK_RUNLEVEL_LUA";				//Tasks will be run with the least privileges.
		case 1: return "TASK_RUNLEVEL_HIGHEST";			//Tasks will be run with the highest privileges.
	}
	return ""; 
}
string getLogonTypeString (int n)
{
	switch (n)
	{
		case 0: return "TASK_LOGON_NONE";
		case 1: return "TASK_LOGON_PASSWORD"; 
		case 2: return "TASK_LOGON_S4U"; 
		case 3: return "TASK_LOGON_INTERACTIVE_TOKEN";
		case 4: return "TASK_LOGON_GROUP"; 
		case 5: return "TASK_LOGON_SERVICE_ACCOUNT"; 
		case 6: return "TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD";
	}
	return ""; 
}
string getState(int n)
{
	switch (n) 
	{
		case 1: return  "TASK_STATE_DISABLED"; 
		case 2: return  "TASK_STATE_QUEUED";
		case 3: return  "TASK_STATE_READY"; 
		case 4: return  "TASK_STATE_RUNNING"; 
		default: return "TASK_STATE_UNKNOWN";
	}
	return ""; 
}
string getTypeTriggerString (TASK_TRIGGER_TYPE2  n) 
{
	switch (n)
	{
		case TASK_TRIGGER_EVENT: return "Starts the task when a specific event occurs.";			
		case TASK_TRIGGER_TIME: return "Starts the task at a specific time of day.";	
		case TASK_TRIGGER_DAILY: return "Starts the task daily.";			
		case TASK_TRIGGER_WEEKLY: return "Starts the task weekly.";	
		case TASK_TRIGGER_MONTHLY: return "Starts the task monthly.";			
		case TASK_TRIGGER_MONTHLYDOW: return "Starts the task every month on a specific day of the week.";	
		case TASK_TRIGGER_IDLE: return "Starts the task when the computer goes into an idle state.";			
		case TASK_TRIGGER_REGISTRATION: return "Starts the task when the task is registered.";	
		case TASK_TRIGGER_BOOT: return "Starts the task when the computer boots.";			
		case TASK_TRIGGER_LOGON: return "Starts the task when a specific user logs on.";	
		case TASK_TRIGGER_SESSION_STATE_CHANGE: return "Triggers the task when a specific session state changes.";			
	}
	return ""; 
}