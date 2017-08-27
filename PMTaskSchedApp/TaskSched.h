#pragma once

#include <Windows.h>
#include <taskschd.h>

#ifdef UNITTEST
// unit test version
// if the location does not match the intended fail point, make the real call
// if the location matches the intended origin, return E_FAIL to simulate failure
#define TESTFAIL(curLoc, call) ((curLoc)==uFailAt?E_FAIL:(call))
#else
// production version makes real call always
#define TESTFAIL(curLoc, call) (call)
#endif

class TaskSched
{
private:
	ITaskService *pTaskService;
	ITaskFolder *pTaskFolder;
	ITaskDefinition *pTaskDefinition;
	IRegistrationInfo *pRegistrationInfo;
	ITriggerCollection *pTriggerCollection;
	ITrigger *pTrigger;
	IAction *pAction;
	IExecAction *pExecAction;
	IRegisteredTask *pRegisteredTask;
	IActionCollection *pActionCollection;
	IPrincipal *pPrincipal;
	ITaskSettings  *pTaskSettings;
	ITimeTrigger *pTimeTrigger;
	IIdleSettings *pIdleSettings;
	
#ifdef UNITTEST
	unsigned uFailAt;
#endif
	unsigned uFailOrigin;
	bool bCoInitDone;
	bool bIsValid;

public:
	TaskSched();
	~TaskSched();

	HRESULT init();
	void term();

	HRESULT CreateTask(CString sExecutable, CString sArgs, CString sTaskName, unsigned uWhen);

	bool IsValid()
	{
		return bIsValid;
	}
	
#ifdef UNITTEST
	void injectError(unsigned uLocation)
	{
		uFailAt = uLocation;
	}
#endif

	void setFailOrigin(unsigned uLocation)
	{
		uFailOrigin = uLocation;
	}
	
	unsigned getFailOrigin() {
		return uFailOrigin;
	}
		
	enum {
		FAIL_NoError = 0,
		FAIL_StartCom,
		FAIL_CoCreate,
		FAIL_CoInitSecurity,
		FAIL_InvalidObject,
		FAIL_Connect,
		FAIL_GetFolder,
		FAIL_DeleteTask,
		FAIL_NewTask,
		FAIL_GetRegistrationInfo,		
		FAIL_PutAuthor,
		FAIL_GetPrincipal,
		FAIL_PutLogonType,
		FAIL_GetSettings,
		FAIL_PutStartWhenAvailable,
		FAIL_GetIdleSettings,
		FAIL_PutWaitTimeout,
		FAIL_GetTriggers,	
		FAIL_CreateTrigger,
		FAIL_QueryInterfaceTT,
		FAIL_PutId,		
		FAIL_PutStartBoundary,
		FAIL_PutEndBoundary,
		FAIL_GetActions,
		FAIL_CreateAction,
		FAIL_QueryInterfaceEA,
		FAIL_PutPath,
		FAIL_PutArguments,
		FAIL_RegisterTask,
	} CaseNumber;

};

