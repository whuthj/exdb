#ifndef DATA_BASE_ENGINE_HEAD_FILE
#define DATA_BASE_ENGINE_HEAD_FILE

#pragma once

#include "Stdafx.h"
#include <ICrsint.h>
#include <math.h>
#include "IUnknownEx.h"
#include "Template.h"

//ADO �����
#import "MSADO15.DLL" rename_namespace("ADOCG") rename("EOF","EndOfFile")
using namespace ADOCG;

//ģ�鶨��
#ifdef _DEBUG
#define XDB_DLL_NAME	TEXT("XDB.dll")			//��� DLL ����
#else
#define XDB_DLL_NAME	TEXT("XDB.dll")			//��� DLL ����
#endif

//COM ��������
typedef _com_error					CComError;							//COM ����

//////////////////////////////////////////////////////////////////////////
//ö�ٶ���

//���ݿ�����
enum DataBaseType
{
    MSSQLSERVER                    =0,                                 //MSSqlserver ���ݿ�
	ORACLE                         =1,                                 //Oracle ���ݿ�
};

//���ݿ�������
enum enADOErrorType
{
	ErrorType_Nothing				=0,									//û�д���
	ErrorType_Connect				=1,									//���Ӵ���
	ErrorType_Other					=2,									//��������
};

//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////

#define VER_IADOError INTERFACE_VERSION(1,1)
static const GUID IID_IADOError={0x66463b5a,0x390c,0x42f9,0x85,0x19,0x13,0x31,0x39,0x36,0xfe,0x8f};

//���ݿ����ӿ�
interface IADOError : public IUnknownEx
{
	//��������
	virtual LPCTSTR __cdecl GetErrorDescribe()=NULL;
	//��������
	virtual enADOErrorType __cdecl GetErrorType()=NULL;
};

//////////////////////////////////////////////////////////////////////////

#define VER_IXDB INTERFACE_VERSION(1,1)
static const GUID IID_IXDB={0xc02b1038,0xfaec,0x42cf,0x0091,0x76,0x64,0x13,0xef,0x64,0xcf,0x99};

//���ݿ����ӽӿ�
interface IXDB : public IUnknownEx
{
	//������
	virtual bool __cdecl OpenConnection()=NULL;
	//�رռ�¼
	virtual bool __cdecl CloseRecordset()=NULL;
	//�ر�����
	virtual bool __cdecl CloseConnection()=NULL;
	//��������
	virtual bool __cdecl TryConnectAgain(bool bFocusConnect, CComError * pComError)=NULL;
	//������Ϣ
	virtual bool __cdecl SetConnectionInfo(DataBaseType type,LPCTSTR szIP, WORD wPort, LPCTSTR szData, LPCTSTR szName, LPCTSTR szPass)=NULL;
	//�Ƿ����Ӵ���
	virtual bool __cdecl IsConnectError()=NULL;
	//�Ƿ��
	virtual bool __cdecl IsRecordsetOpened()=NULL;
	//�����ƶ�
	virtual void __cdecl MoveToNext()=NULL;
	//�Ƶ���ͷ
	virtual void __cdecl MoveToFirst()=NULL;
	//�Ƿ����
	virtual bool __cdecl IsEndRecordset()=NULL;
	//��ȡ��Ŀ
	virtual long __cdecl GetRecordCount()=NULL;
	//��ȡ��С
	virtual long __cdecl GetActualSize(LPCTSTR pszParamName)=NULL;
	//�󶨶���
	virtual bool __cdecl BindToRecordset(CADORecordBinding * pBind)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, BYTE & bValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, WORD & wValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, INT & nValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, LONG & lValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DWORD & ulValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, UINT & ulValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DOUBLE & dbValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, __int64 & llValue)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, TCHAR szBuffer[], UINT uSize)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, COleDateTime & Time)=NULL;
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, bool & bValue)=NULL;
	//���ô洢����
	virtual void __cdecl SetSPName(LPCTSTR pszSpName)=NULL;
	//�������
	virtual void __cdecl AddParamter(LPCTSTR pszName, ADOCG::ParameterDirectionEnum Direction, ADOCG::DataTypeEnum Type, long lSize, _variant_t & vtValue)=NULL;
	//ɾ������
	virtual void __cdecl ClearAllParameters()=NULL;
	//��ò���
	virtual void __cdecl GetParameterValue(LPCTSTR pszParamName, _variant_t & vtValue)=NULL;
	//��ȡ������ֵ
	virtual long __cdecl GetReturnValue()=NULL;
	//ִ�����
	virtual bool __cdecl Execute(LPCTSTR pszCommand,bool bRecordset)=NULL;
	//ִ������
	virtual bool __cdecl ExecuteCommand(bool bRecordset)=NULL;
};

//////////////////////////////////////////////////////////////////////////

//ADO ������
class CADOError:IADOError
{
	//��������
protected:
	enADOErrorType					m_enErrorType;						//�������
	TCHAR							m_strErrorDescribe[50];					//������Ϣ

	//��������
public:
	//���캯��
	CADOError();
	//��������
	virtual ~CADOError();

	//�����ӿ�
public:
	//�ͷŶ���
	virtual bool __cdecl Release() { return true; }
	//�Ƿ���Ч
	virtual bool __cdecl IsValid() { return AfxIsValidAddress(this,sizeof(CADOError))?true:false; }
	//�ӿڲ�ѯ
	virtual void * __cdecl QueryInterface(const IID & Guid, DWORD dwQueryVer);

	//���ܽӿ�
public:
	//��������
	virtual enADOErrorType __cdecl GetErrorType() { return m_enErrorType; }
	//��������
	virtual LPCTSTR __cdecl GetErrorDescribe() { return m_strErrorDescribe; }
	
	//���ܺ���
public:
	//���ô���
	void SetErrorInfo(enADOErrorType enErrorType, LPCTSTR pszDescribe);
};


//////////////////////////////////////////////////////////////////////////

//���ݿ����
class CXDB : public IXDB
{
	//��Ϣ����
protected:
	CADOError						m_ADOError;							//�������
	TCHAR							m_strConnect[100];						//�����ַ���
	TCHAR							m_strErrorDescribe[100];					//������Ϣ

	//״̬����
protected:
	DWORD							m_dwConnectCount;					//���Դ���
	DWORD							m_dwConnectErrorTime;				//����ʱ��
	const DWORD						m_dwResumeConnectCount;				//�ָ�����
	const DWORD						m_dwResumeConnectTime;				//�ָ�ʱ��

	//�ں˱���
protected:
	_CommandPtr						m_DBCommand;						//�������
	_RecordsetPtr					m_DBRecordset;						//��¼������
	_ConnectionPtr					m_DBConnection;						//���ݿ����

	//��������
public:
	//���캯��
	CXDB();
	//��������
	virtual ~CXDB();

	//�����ӿ�
public:
	//�ͷŶ���
	virtual bool __cdecl Release() { if (IsValid()) delete this; return true; }
	//�Ƿ���Ч
	virtual bool __cdecl IsValid() { return AfxIsValidAddress(this,sizeof(CXDB))?true:false; }
	//�ӿڲ�ѯ
	virtual void * __cdecl QueryInterface(const IID & Guid, DWORD dwQueryVer);

	//����ӿ�
public:
	//������
	virtual bool __cdecl OpenConnection();
	//�رռ�¼
	virtual bool __cdecl CloseRecordset();
	//�ر�����
	virtual bool __cdecl CloseConnection();
	//��������
	virtual bool __cdecl TryConnectAgain(bool bFocusConnect, CComError * pComError);
	//������Ϣ
	virtual bool __cdecl SetConnectionInfo(DataBaseType type,LPCTSTR szIP, WORD wPort, LPCTSTR szData, LPCTSTR szName, LPCTSTR szPass);

	//״̬�ӿ�
public:
	//�Ƿ����Ӵ���
	virtual bool __cdecl IsConnectError();
	//�Ƿ��
	virtual bool __cdecl IsRecordsetOpened();

	//��¼���ӿ�
public:
	//�����ƶ�
	virtual void __cdecl MoveToNext();
	//�Ƶ���ͷ
	virtual void __cdecl MoveToFirst();
	//�Ƿ����
	virtual bool __cdecl IsEndRecordset();
	//��ȡ��Ŀ
	virtual long __cdecl GetRecordCount();
	//��ȡ��С
	virtual long __cdecl GetActualSize(LPCTSTR pszParamName);
	//�󶨶���
	virtual bool __cdecl BindToRecordset(CADORecordBinding * pBind);

	//�ֶνӿ�
public:
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, BYTE & bValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, WORD & wValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, INT & nValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, LONG & lValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DWORD & ulValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, UINT & ulValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DOUBLE & dbValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, __int64 & llValue);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, TCHAR szBuffer[], UINT uSize);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, COleDateTime & Time);
	//��ȡ����
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, bool & bValue);

	//�������ӿ�
public:
	//���ô洢����
	virtual void __cdecl SetSPName(LPCTSTR pszSpName);
	//�������
	virtual void __cdecl AddParamter(LPCTSTR pszName, ADOCG::ParameterDirectionEnum Direction, ADOCG::DataTypeEnum Type, long lSize, _variant_t & vtValue);
	//ɾ������
	virtual void __cdecl ClearAllParameters();
	//��ò���
	virtual void __cdecl GetParameterValue(LPCTSTR pszParamName, _variant_t & vtValue);
	//��ȡ������ֵ
	virtual long __cdecl GetReturnValue();

	//ִ�нӿ�
public:
	//ִ�����
	virtual bool __cdecl Execute(LPCTSTR pszCommand,bool bRecordset);
	//ִ������
	virtual bool __cdecl ExecuteCommand(bool bRecordset);

	//�ڲ�����
private:
	//��ȡ����
	LPCTSTR GetComErrorDescribe(CComError & ComError);
	//���ô���
	void SetErrorInfo(enADOErrorType enErrorType, LPCTSTR pszDescribe);
};

//////////////////////////////////////////////////////////////////////////

//���ݿ����渨����
class CXDBHelper : public CTempldateHelper<IXDB>
{
	//��������
public:
	//���캯��
	CXDBHelper(void) : CTempldateHelper<IXDB>(IID_IXDB,
		VER_IXDB,XDB_DLL_NAME,TEXT("CreateXDB")) { }
};

//////////////////////////////////////////////////////////////////////////

#endif