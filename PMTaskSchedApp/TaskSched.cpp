//-----------------------------------------------------------
// Use COM to create a task using the Windows task scheduler
//
// Pat Mancuso - 8/25/2017
//
//-----------------------------------------------------------


#include "stdafx.h"

#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>
#include <taskschd.h>
#include <atltime.h>

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")

#include "TaskSched.h"

TaskSched::TaskSched()
{
	// real work is done in init() to allow constructor to be called
	// so that fault injection can be set up prior to executing that code
}

TaskSched::~TaskSched()
{
	// real work is done in term so that error cases can clean up
	// immediately by calling term.  Destructor calls term to 
	// handle cases without prior errors.
	term();
}

// Set up a connection to the task scheduler
HRESULT TaskSched::init()
{
	HRESULT hrResult;
	
	bCoInitDone = false;
	bIsValid = false;

	// Do as much as we can ahead of the actual CreateTask call

	//  Start up COM
	hrResult = TESTFAIL(FAIL_StartCom, CoInitializeEx(NULL, COINIT_MULTITHREADED));
	if (FAILED(hrResult)) 
	{
		setFailOrigin(FAIL_StartCom);
	}
	else
	{
		bCoInitDone = true;

		// Create/find the task scheduler
		hrResult = TESTFAIL(FAIL_CoCreate, CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pTaskService));
		if (FAILED(hrResult))
		{
			setFailOrigin(FAIL_CoCreate);
		}
		else
		{
			// Establish a simple security configuration
			hrResult = TESTFAIL(FAIL_CoInitSecurity, CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_IMPERSONATE,	NULL, 0, NULL));
			if (FAILED(hrResult))
			{
				setFailOrigin(FAIL_CoInitSecurity);
			}
			else
			{
				// Made it this far...mark it valid
				TESTFAIL(FAIL_InvalidObject, bIsValid = true);
			}
		}
	}

	if (bIsValid == false)
	{
		term();
	}

	return hrResult;
}

// Define local exception class to use throw/catch and avoid ugly huge nested if/then/else 
class TSFailure : std::exception
{
	unsigned uOrigin;

public:
	// Make note of the case where the exception occurred 
	TSFailure(unsigned uLocation)
	{
		uOrigin = uLocation;
	}

	unsigned getOrigin()
	{
		return uOrigin;
	}
};

// Create the actual task using the task scheduler
//
// Arguments:
//
// sExecutable		IN	full path to executable to be run
// sArgs			IN	arguments to be passed to executable
// sTaskname		IN	name for task in scheduler
// uWhen			IN	number of seconds in the future to initiate the task
//
HRESULT TaskSched::CreateTask(CString sExecutable, CString sArgs, CString sTaskName, unsigned uWhen)
{
	HRESULT hrResult = E_FAIL;

	if (bIsValid == false)
	{
		// Make sure the init worked...
		setFailOrigin(FAIL_InvalidObject);
	}
	else
	{
		// we're ready to talk to the service...
		try {

			// do some basic arg validation
			if (sExecutable.GetLength() == 0 ||
				sTaskName.GetLength() == 0 ||
				uWhen == 0)
			{
				// invalid args
				hrResult = E_INVALIDARG;
				setFailOrigin(FAIL_InvalidArgs);
				throw TSFailure(FAIL_InvalidArgs);
			}

			//  Connect to the task service.
			hrResult = TESTFAIL(FAIL_Connect, pTaskService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t()));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_Connect);

			// Get folder to store task
			hrResult = TESTFAIL(FAIL_GetFolder, pTaskService->GetFolder(L"\\", &pTaskFolder));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_GetFolder);

			// Delete any pre-existing task of same name
			hrResult = TESTFAIL(FAIL_DeleteTask, pTaskFolder->DeleteTask(_bstr_t(sTaskName), 0));
			// note special case here - not found is acceptable
			if ((hrResult != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) && FAILED(hrResult))
				throw TSFailure(FAIL_DeleteTask);

			//  Make a new task
			hrResult = TESTFAIL(FAIL_NewTask, pTaskService->NewTask(0, &pTaskDefinition));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_NewTask);

			//  Get registration info
			hrResult = TESTFAIL(FAIL_GetRegistrationInfo, pTaskDefinition->get_RegistrationInfo(&pRegistrationInfo));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_GetRegistrationInfo);

			// Set author
			hrResult = TESTFAIL(FAIL_PutAuthor, pRegistrationInfo->put_Author(L"TaskSched"));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutAuthor);

			// Get principal
			hrResult = TESTFAIL(FAIL_GetPrincipal, pTaskDefinition->get_Principal(&pPrincipal));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_GetPrincipal);

			// Set logon type
			hrResult = TESTFAIL(FAIL_PutLogonType, pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutLogonType);

			// Get settings for task
			hrResult = TESTFAIL(FAIL_GetSettings, pTaskDefinition->get_Settings(&pTaskSettings));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_GetSettings);

			// Task can start during time window
			hrResult = TESTFAIL(FAIL_PutStartWhenAvailable, pTaskSettings->put_StartWhenAvailable(VARIANT_TRUE));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutStartWhenAvailable);

			// Get Idle Settings
			hrResult = TESTFAIL(FAIL_GetIdleSettings, pTaskSettings->get_IdleSettings(&pIdleSettings));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_GetIdleSettings);

			// Set timeout for idle to be 5 minutes
			hrResult = TESTFAIL(FAIL_PutWaitTimeout, pIdleSettings->put_WaitTimeout(L"PT5M"));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutWaitTimeout);

			// Get triggers
			hrResult = TESTFAIL(FAIL_GetTriggers, pTaskDefinition->get_Triggers(&pTriggerCollection));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_GetTriggers);

			// Create new trigger
			hrResult = TESTFAIL(FAIL_CreateTrigger, pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_CreateTrigger);

			// Get interface for time triggers
			hrResult = TESTFAIL(FAIL_QueryInterfaceTT, pTrigger->QueryInterface(IID_ITimeTrigger, (void**)&pTimeTrigger));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_QueryInterfaceTT);

			// Set ID of time trigger
			hrResult = TESTFAIL(FAIL_PutId, pTimeTrigger->put_Id(_bstr_t(L"PMTaskSchedTrigger")));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutId);

			// Use the uWhen value with the current time to calculate a start time relative to 'now'
			CTime tStart = CTime::GetCurrentTime();
			CTime tEnd;
			// create delta from specified value
			CTimeSpan tsDelay(0, 0, 0, uWhen);
			// create 10 minute end of window
			CTimeSpan tsTimeout(0, 0, 10, 0);

			// Start time is current time plus specified delay
			tStart += tsDelay;

			// End time is start plus timeout
			tEnd = tStart;
			tEnd += tsTimeout;

			// Set start and end boundaries for time schedule
			hrResult = TESTFAIL(FAIL_PutStartBoundary, pTimeTrigger->put_StartBoundary(_bstr_t(tStart.FormatGmt(L"%Y-%m-%dT%H:%M:%SZ"))));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutStartBoundary);

			hrResult = TESTFAIL(FAIL_PutEndBoundary, pTimeTrigger->put_EndBoundary(_bstr_t(tEnd.FormatGmt(L"%Y-%m-%dT%H:%M:%SZ"))));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutEndBoundary);

			//  Get the task action collection pointer.
			hrResult = TESTFAIL(FAIL_GetActions, pTaskDefinition->get_Actions(&pActionCollection));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_GetActions);

			// Create action
			hrResult = TESTFAIL(FAIL_CreateAction, pActionCollection->Create(TASK_ACTION_EXEC, &pAction));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_CreateAction);

			//  Get executable action
			hrResult = TESTFAIL(FAIL_QueryInterfaceEA, pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_QueryInterfaceEA);

			//  Set the path of the executable to notepad.exe.
			hrResult = TESTFAIL(FAIL_PutPath, pExecAction->put_Path(_bstr_t(sExecutable)));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutPath);

			//  Set Arguments
			hrResult = TESTFAIL(FAIL_PutArguments, pExecAction->put_Arguments(_bstr_t(sArgs)));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_PutArguments);

			// last step...save task
			hrResult = TESTFAIL(FAIL_RegisterTask, pTaskFolder->RegisterTaskDefinition(_bstr_t(sTaskName), pTaskDefinition, TASK_CREATE_OR_UPDATE, _variant_t(), _variant_t(), TASK_LOGON_INTERACTIVE_TOKEN, _variant_t(L""), &pRegisteredTask));
			if (FAILED(hrResult))
				throw TSFailure(FAIL_RegisterTask);

			// it worked!!!
			setFailOrigin(FAIL_NoError);
		}
		catch (TSFailure& TSFException)
		{
			setFailOrigin(TSFException.getOrigin());
		}
	}


	// If something went wrong...clean up to ensure nothing is left half alive
	if (FAILED(hrResult))
		term();

	return hrResult;
}

// Tear down all the things that were built
void TaskSched::term()
{
	if (pTaskService)
	{
		pTaskService->Release();
		pTaskService = 0;
	}

	if (pTaskFolder)
	{
		pTaskFolder->Release();
		pTaskFolder = 0;
	}

	if (pTaskDefinition)
	{
		pTaskDefinition->Release();
		pTaskDefinition = 0;
	}

	if (pRegistrationInfo)
	{
		pRegistrationInfo->Release();
		pRegistrationInfo = 0;
	}

	if (pTriggerCollection)
	{
		pTriggerCollection->Release();
		pTriggerCollection = 0;
	}

	if (pTrigger)
	{
		pTrigger->Release();
		pTrigger = 0;
	}

	if (pAction)
	{
		pAction->Release();
		pAction = 0;
	}

	if (pExecAction)
	{
		pExecAction->Release();
		pExecAction = 0;
	}

	if (pRegisteredTask)
	{
		pRegisteredTask->Release();
		pRegisteredTask = 0;
	}

	if (pActionCollection)
	{
		pActionCollection->Release();
		pActionCollection = 0;
	}
	
	if (pPrincipal)
	{
		pPrincipal->Release();
		pPrincipal = 0;
	}
	
	if (pTaskSettings)
	{
		pTaskSettings->Release();
		pTaskSettings = 0;
	}

	if (pIdleSettings)
	{
		pIdleSettings->Release();
		pIdleSettings = 0;
	}

	if (bCoInitDone == true)
	{
		CoUninitialize();
		bCoInitDone = false;
	}

	bIsValid = false;
}



