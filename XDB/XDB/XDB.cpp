#include "XDB.h"

//�궨��
_COM_SMARTPTR_TYPEDEF(IADORecordBinding,__uuidof(IADORecordBinding));

//Ч������
#define EfficacyResult(hResult) { if (FAILED(hResult)) _com_issue_error(hResult); }

//////////////////////////////////////////////////////////////////////////

//���캯��
CADOError::CADOError()
{
	m_enErrorType=ErrorType_Nothing;
}

//��������
CADOError::~CADOError()
{
}

//�ӿڲ�ѯ
void * __cdecl CADOError::QueryInterface(const IID & Guid, DWORD dwQueryVer)
{
	QUERYINTERFACE(IADOError,Guid,dwQueryVer);
	QUERYINTERFACE_IUNKNOWNEX(IADOError,Guid,dwQueryVer);
	return NULL;
}

//���ô���
void CADOError::SetErrorInfo(enADOErrorType enErrorType, LPCTSTR pszDescribe)
{
	//���ô���
	m_enErrorType=enErrorType;
	_snprintf(m_strErrorDescribe,sizeof(m_strErrorDescribe),TEXT("%s"),pszDescribe);

	return;
}

//////////////////////////////////////////////////////////////////////////

//���캯��
CXDB::CXDB() : m_dwResumeConnectCount(30L),m_dwResumeConnectTime(30L)
{
	//״̬����
	m_dwConnectCount=0;
	m_dwConnectErrorTime=0L;

	//��������
	m_DBCommand.CreateInstance(__uuidof(Command));
	m_DBRecordset.CreateInstance(__uuidof(Recordset));
	m_DBConnection.CreateInstance(__uuidof(Connection));

	if (m_DBCommand==NULL) throw TEXT("���ݿ�������󴴽�ʧ��");
	if (m_DBRecordset==NULL) throw TEXT("���ݿ��¼�����󴴽�ʧ��");
	if (m_DBConnection==NULL) throw TEXT("���ݿ����Ӷ��󴴽�ʧ��");

	//���ñ���
	m_DBCommand->CommandType=adCmdStoredProc;

	return;
}

//��������
CXDB::~CXDB()
{
	//�ر�����
	CloseConnection();

	//�ͷŶ���
	m_DBCommand.Release();
	m_DBRecordset.Release();
	m_DBConnection.Release();

	return;
}

//�ӿڲ�ѯ
void * __cdecl CXDB::QueryInterface(const IID & Guid, DWORD dwQueryVer)
{
	QUERYINTERFACE(IXDB,Guid,dwQueryVer);
	QUERYINTERFACE_IUNKNOWNEX(IXDB,Guid,dwQueryVer);
	return NULL;
}


//������
bool __cdecl CXDB::OpenConnection()
{
	//�������ݿ�
	try
	{
		//�ر�����
		CloseConnection();

		//�������ݿ�
		EfficacyResult(m_DBConnection->Open(_bstr_t(m_strConnect),L"",L"",adConnectUnspecified));
		m_DBConnection->CursorLocation=adUseClient;
		m_DBCommand->ActiveConnection=m_DBConnection;

		//���ñ���
		m_dwConnectCount=0L;
		m_dwConnectErrorTime=0L;

		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//�رռ�¼
bool __cdecl CXDB::CloseRecordset()
{
	try 
	{
		if (IsRecordsetOpened()) EfficacyResult(m_DBRecordset->Close());
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//�ر�����
bool __cdecl CXDB::CloseConnection()
{
	try
	{
		CloseRecordset();
		if ((m_DBConnection!=NULL)&&(m_DBConnection->GetState()!=adStateClosed))
		{
			EfficacyResult(m_DBConnection->Close());
		}
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��������
bool __cdecl CXDB::TryConnectAgain(bool bFocusConnect, CComError * pComError)
{
	try
	{
		//�ж�����
		bool bReConnect=bFocusConnect;
		if (bReConnect==false)
		{
			DWORD dwNowTime=(DWORD)time(NULL);
			if ((m_dwConnectErrorTime+m_dwResumeConnectTime)>dwNowTime) bReConnect=true;
		}
		if ((bReConnect==false)&&(m_dwConnectCount>m_dwResumeConnectCount)) bReConnect=true;

		//���ñ���
		m_dwConnectCount++;
		m_dwConnectErrorTime=(DWORD)time(NULL);
		if (bReConnect==false)
		{
			if (pComError!=NULL) SetErrorInfo(ErrorType_Connect,GetComErrorDescribe(*pComError));
			return false;
		}

		//��������
		OpenConnection();
		return true;
	}
	catch (...)
	{
		//�������Ӵ���
		if (pComError!=NULL) SetErrorInfo(ErrorType_Connect,GetComErrorDescribe(*pComError));
	}

	return false;
}

//������Ϣ
bool __cdecl CXDB::SetConnectionInfo(DataBaseType type,LPCTSTR szIP, WORD wPort, LPCTSTR szData, LPCTSTR szName, LPCTSTR szPass)
{
	//���������ַ���
	if(type==MSSQLSERVER)
	{
		_snprintf(m_strConnect,sizeof(m_strConnect),TEXT("Provider=SQLOLEDB.1;Password=%s;Persist Security Info=True;User ID=%s;Initial Catalog=%s;Data Source=%s,%ld;"),
		szPass,szName,szData,szIP,wPort);
	}
	else if(type==ORACLE)
	{
	    _snprintf(m_strConnect,sizeof(m_strConnect),TEXT("Provider=OraOLEDB.Oracle;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True;"),
		szPass,szName,szData);
	}
	
	return true;
}

//�Ƿ����Ӵ���
bool __cdecl CXDB::IsConnectError()
{
	try 
	{
		//״̬�ж�
		if (m_DBConnection==NULL) return true;
		if (m_DBConnection->GetState()==adStateClosed) return true;

		//�����ж�
		long lErrorCount=m_DBConnection->Errors->Count;
		if (lErrorCount>0L)
		{
	        ErrorPtr pError=NULL;
			for(long i=0;i<lErrorCount;i++)
			{
				pError=m_DBConnection->Errors->GetItem(i);
				if (pError->Number==0x80004005) return true;
			}
		}

		return false;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//�Ƿ��
bool __cdecl CXDB::IsRecordsetOpened()
{
	if (m_DBRecordset==NULL) return false;
	if (m_DBRecordset->GetState()==adStateClosed) return false;
	return true;
}

//�����ƶ�
void __cdecl CXDB::MoveToNext()
{
	try 
	{ 
		m_DBRecordset->MoveNext(); 
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return;
}

//�Ƶ���ͷ
void __cdecl CXDB::MoveToFirst()
{
	try 
	{ 
		m_DBRecordset->MoveFirst(); 
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return;
}

//�Ƿ����
bool __cdecl CXDB::IsEndRecordset()
{
	try 
	{
		return (m_DBRecordset->EndOfFile==VARIANT_TRUE); 
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return true;
}

//��ȡ��Ŀ
long __cdecl CXDB::GetRecordCount()
{
	try
	{
		if (m_DBRecordset==NULL) return 0;
		return m_DBRecordset->GetRecordCount();
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return 0;
}

//��ȡ��С
long __cdecl CXDB::GetActualSize(LPCTSTR pszParamName)
{
	ASSERT(pszParamName!=NULL);
	try 
	{
		return m_DBRecordset->Fields->Item[pszParamName]->ActualSize;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return -1;
}

//�󶨶���
bool __cdecl CXDB::BindToRecordset(CADORecordBinding * pBind)
{
	ASSERT(pBind!=NULL);
	try 
	{
        IADORecordBindingPtr pIBind(m_DBRecordset);
		pIBind->BindToRecordset(pBind);
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, BYTE & bValue)
{
	try
	{
		bValue=0;
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		switch(vtFld.vt)
		{
		case VT_BOOL:
			{
				bValue=(vtFld.boolVal!=0)?1:0;
				break;
			}
		case VT_I2:
		case VT_UI1:
			{
				bValue=(vtFld.iVal>0)?1:0;
				break;
			}
		case VT_NULL:
		case VT_EMPTY:
			{
				bValue=0;
				break;
			}
		default: bValue=(BYTE)vtFld.iVal;
		}	
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, UINT & ulValue)
{
	try
	{
		ulValue=0L;
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		if ((vtFld.vt!=VT_NULL)&&(vtFld.vt!=VT_EMPTY)) ulValue=vtFld.lVal;
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, DOUBLE & dbValue)
{	
	try
	{
		dbValue=0.0L;
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		switch(vtFld.vt)
		{
		case VT_R4:
			{
				dbValue=vtFld.fltVal;
				break;
			}
		case VT_R8:
			{
				dbValue=vtFld.dblVal;
				break;
			}
		case VT_DECIMAL:
			{
				dbValue=vtFld.decVal.Lo32;
				dbValue*=(vtFld.decVal.sign==128)?-1:1;
				dbValue/=pow(10,vtFld.decVal.scale); 
				break;
			}
		case VT_UI1:
			{
				dbValue=vtFld.iVal;
				break;
			}
		case VT_I2:
		case VT_I4:
			{
				dbValue=vtFld.lVal;
				break;
			}
		case VT_NULL:
		case VT_EMPTY:
			{
				dbValue=0.0L;
				break;
			}
		default: dbValue=vtFld.dblVal;
		}
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, LONG & lValue)
{
	try
	{
		lValue=0L;
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		if ((vtFld.vt!=VT_NULL)&&(vtFld.vt!=VT_EMPTY)) lValue=vtFld.lVal;
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, DWORD & ulValue)
{
	try
	{
		ulValue=0L;
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		if ((vtFld.vt!=VT_NULL)&&(vtFld.vt!=VT_EMPTY)) ulValue=vtFld.ulVal;
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, INT & nValue)
{
	try
	{
		nValue=0;
		_variant_t vtFld = m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		switch(vtFld.vt)
		{
		case VT_BOOL:
			{
				nValue = vtFld.boolVal;
				break;
			}
		case VT_I2:
		case VT_UI1:
			{
				nValue = vtFld.iVal;
				break;
			}
		case VT_NULL:
		case VT_EMPTY:
			{
				nValue = 0;
				break;
			}
		default: nValue = vtFld.iVal;
		}	
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, __int64 & llValue)
{
	try
	{
		llValue=0L;
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		if ((vtFld.vt!=VT_NULL)&&(vtFld.vt!=VT_EMPTY)) 
		{
			llValue=vtFld.lVal;
		}
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, TCHAR szBuffer[], UINT uSize)
{
	try
	{
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		if (vtFld.vt==VT_BSTR)
		{
			lstrcpy(szBuffer,(char*)_bstr_t(vtFld));
			return true;
		}
		return false;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, WORD & wValue)
{
	try
	{
		wValue=0L;
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		if ((vtFld.vt!=VT_NULL)&&(vtFld.vt!=VT_EMPTY)) wValue=(WORD)vtFld.ulVal;
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, COleDateTime & Time)
{
	try
	{
		_variant_t vtFld = m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		switch(vtFld.vt) 
		{
		case VT_DATE:
			{
				COleDateTime TempTime(vtFld);
				Time=TempTime;
				break;
			}
		case VT_EMPTY:
		case VT_NULL:
			{
				Time.SetStatus(COleDateTime::null);
				break;
			}
		default: return false;
		}
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ����
bool __cdecl CXDB::GetFieldValue(LPCTSTR lpFieldName, bool & bValue)
{
	try
	{
		_variant_t vtFld=m_DBRecordset->Fields->GetItem(lpFieldName)->Value;
		switch(vtFld.vt) 
		{
		case VT_BOOL:
			{
				bValue=(vtFld.boolVal==0)?false:true;
				break;
			}
		case VT_EMPTY:
		case VT_NULL:
			{
				bValue = false;
				break;
			}
		default:return false;
		}
		return true;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return false;
}

//��ȡ������ֵ
long __cdecl CXDB::GetReturnValue()
{
	try 
	{
        _ParameterPtr Parameter;
		long lParameterCount=m_DBCommand->Parameters->Count;
		for (long i=0;i<lParameterCount;i++)
		{
			Parameter=m_DBCommand->Parameters->Item[i];
			if (Parameter->Direction==adParamReturnValue) return Parameter->Value.lVal;
		}
		ASSERT(FALSE);
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return 0;
}

//ɾ������
void __cdecl CXDB::ClearAllParameters()
{
	try 
	{
		long lParameterCount=m_DBCommand->Parameters->Count;
		long temp=0L;
		if (lParameterCount>0L)
		{
			for (long i=lParameterCount;i>0;i--)
			{
				temp=i-1;
				m_DBCommand->Parameters->Delete((_variant_t &)temp);
			}
		}
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return;
}

//���ô洢����
void __cdecl CXDB::SetSPName(LPCTSTR pszSpName)
{
	ASSERT(pszSpName!=NULL);
	try 
	{ 
		m_DBCommand->CommandText=pszSpName; 
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return;
}

//��ò���
void __cdecl CXDB::GetParameterValue(LPCTSTR pszParamName, _variant_t & vtValue)
{
	//Ч�����
	ASSERT(pszParamName!=NULL);

	//��ȡ����
	try 
	{
		vtValue.Clear();
		vtValue=m_DBCommand->Parameters->Item[pszParamName]->Value;
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return;
}

//�������
void __cdecl CXDB::AddParamter(LPCTSTR pszName, ADOCG::ParameterDirectionEnum Direction, ADOCG::DataTypeEnum Type, long lSize, _variant_t & vtValue)
{
	ASSERT(pszName!=NULL);
	try 
	{
        _ParameterPtr Parameter=m_DBCommand->CreateParameter(pszName,Type,Direction,lSize,vtValue);
		m_DBCommand->Parameters->Append(Parameter);
	}
	catch (CComError & ComError) { SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError)); }

	return;
}

//ִ�����
bool __cdecl CXDB::Execute(LPCTSTR pszCommand,bool bRecordset)
{
	ASSERT(pszCommand!=NULL);
	try
	{
		if(bRecordset==true)
		{
		   CloseRecordset();
		   m_DBConnection->CursorLocation=adUseClient;
		   m_DBRecordset=m_DBConnection->Execute(pszCommand,NULL,adCmdText);
		}
		else
		{
		   m_DBConnection->CursorLocation=adUseClient;
		   m_DBConnection->Execute(pszCommand,NULL,adExecuteNoRecords);
		}
		return true;
	}
	catch (CComError & ComError) 
	{ 
		if (IsConnectError()==true)	TryConnectAgain(false,&ComError);
		else SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError));
	}

	return false;
}

//ִ������
bool __cdecl CXDB::ExecuteCommand(bool bRecordset)
{
	try 
	{
		//ִ������
		if (bRecordset==true)
		{
			//�رռ�¼��
		    CloseRecordset();

			m_DBRecordset->PutRefSource(m_DBCommand);
			m_DBRecordset->CursorLocation=adUseClient;
			EfficacyResult(m_DBRecordset->Open((IDispatch *)m_DBCommand,vtMissing,adOpenForwardOnly,adLockReadOnly,adOptionUnspecified));
		}
		else 
		{
			m_DBConnection->CursorLocation=adUseClient;
			EfficacyResult(m_DBCommand->Execute(NULL,NULL,adExecuteNoRecords));
		}
		return true;
	}
	catch (CComError & ComError) 
	{ 
		if (IsConnectError()==true)	TryConnectAgain(false,&ComError);
		else SetErrorInfo(ErrorType_Other,GetComErrorDescribe(ComError));
	}

	return false;
}

//��ȡ����
LPCTSTR CXDB::GetComErrorDescribe(CComError & ComError)
{
	_bstr_t bstrDescribe(ComError.Description());
	_snprintf(m_strErrorDescribe,sizeof(m_strErrorDescribe),TEXT("ADO ����0x%8x��%s"),ComError.Error(),(LPCTSTR)bstrDescribe);
	return m_strErrorDescribe;
}

//���ô���
void CXDB::SetErrorInfo(enADOErrorType enErrorType, LPCTSTR pszDescribe)
{
	m_ADOError.SetErrorInfo(enErrorType,pszDescribe);
	return;
}

//////////////////////////////////////////////////////////////////////////

//DLL ����������
extern "C" int APIENTRY DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	UNREFERENCED_PARAMETER(lpReserved);
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		//��ʼ�� COM
		CoInitialize(NULL);
	}
	else if (dwReason==DLL_PROCESS_DETACH)
	{
		CoUninitialize();
	}
	
	return 1;
}

//////////////////////////////////////////////////////////////////////////

//����������
extern "C" __declspec(dllexport) void * __cdecl CreateXDB(const GUID & Guid, DWORD dwInterfaceVer)
{
	//��������
	CXDB * pDataBase=NULL;
	try
	{
		pDataBase=new CXDB();
		if (pDataBase==NULL) throw TEXT("����ʧ��");
		void * pObject=pDataBase->QueryInterface(Guid,dwInterfaceVer);
		if (pObject==NULL) throw TEXT("�ӿڲ�ѯʧ��");
		return pObject;
	}
	catch (...) {}
	
	//�������
	SafeDelete(pDataBase);
	return NULL;
}

//////////////////////////////////////////////////////////////////////////
