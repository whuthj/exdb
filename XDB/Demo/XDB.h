#ifndef DATA_BASE_ENGINE_HEAD_FILE
#define DATA_BASE_ENGINE_HEAD_FILE

#pragma once

#include "Stdafx.h"
#include <ICrsint.h>
#include <math.h>
#include "IUnknownEx.h"
#include "Template.h"

//ADO 导入库
#import "MSADO15.DLL" rename_namespace("ADOCG") rename("EOF","EndOfFile")
using namespace ADOCG;

//模块定义
#ifdef _DEBUG
#define XDB_DLL_NAME	TEXT("XDB.dll")			//组件 DLL 名字
#else
#define XDB_DLL_NAME	TEXT("XDB.dll")			//组件 DLL 名字
#endif

//COM 错误类型
typedef _com_error					CComError;							//COM 错误

//////////////////////////////////////////////////////////////////////////
//枚举定义

//数据库类型
enum DataBaseType
{
    MSSQLSERVER                    =0,                                 //MSSqlserver 数据库
	ORACLE                         =1,                                 //Oracle 数据库
};

//数据库错误代码
enum enADOErrorType
{
	ErrorType_Nothing				=0,									//没有错误
	ErrorType_Connect				=1,									//连接错误
	ErrorType_Other					=2,									//其他错误
};

//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////

#define VER_IADOError INTERFACE_VERSION(1,1)
static const GUID IID_IADOError={0x66463b5a,0x390c,0x42f9,0x85,0x19,0x13,0x31,0x39,0x36,0xfe,0x8f};

//数据库错误接口
interface IADOError : public IUnknownEx
{
	//错误描述
	virtual LPCTSTR __cdecl GetErrorDescribe()=NULL;
	//错误类型
	virtual enADOErrorType __cdecl GetErrorType()=NULL;
};

//////////////////////////////////////////////////////////////////////////

#define VER_IXDB INTERFACE_VERSION(1,1)
static const GUID IID_IXDB={0xc02b1038,0xfaec,0x42cf,0x0091,0x76,0x64,0x13,0xef,0x64,0xcf,0x99};

//数据库连接接口
interface IXDB : public IUnknownEx
{
	//打开连接
	virtual bool __cdecl OpenConnection()=NULL;
	//关闭记录
	virtual bool __cdecl CloseRecordset()=NULL;
	//关闭连接
	virtual bool __cdecl CloseConnection()=NULL;
	//重新连接
	virtual bool __cdecl TryConnectAgain(bool bFocusConnect, CComError * pComError)=NULL;
	//设置信息
	virtual bool __cdecl SetConnectionInfo(DataBaseType type,LPCTSTR szIP, WORD wPort, LPCTSTR szData, LPCTSTR szName, LPCTSTR szPass)=NULL;
	//是否连接错误
	virtual bool __cdecl IsConnectError()=NULL;
	//是否打开
	virtual bool __cdecl IsRecordsetOpened()=NULL;
	//往下移动
	virtual void __cdecl MoveToNext()=NULL;
	//移到开头
	virtual void __cdecl MoveToFirst()=NULL;
	//是否结束
	virtual bool __cdecl IsEndRecordset()=NULL;
	//获取数目
	virtual long __cdecl GetRecordCount()=NULL;
	//获取大小
	virtual long __cdecl GetActualSize(LPCTSTR pszParamName)=NULL;
	//绑定对象
	virtual bool __cdecl BindToRecordset(CADORecordBinding * pBind)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, BYTE & bValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, WORD & wValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, INT & nValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, LONG & lValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DWORD & ulValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, UINT & ulValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DOUBLE & dbValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, __int64 & llValue)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, TCHAR szBuffer[], UINT uSize)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, COleDateTime & Time)=NULL;
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, bool & bValue)=NULL;
	//设置存储过程
	virtual void __cdecl SetSPName(LPCTSTR pszSpName)=NULL;
	//插入参数
	virtual void __cdecl AddParamter(LPCTSTR pszName, ADOCG::ParameterDirectionEnum Direction, ADOCG::DataTypeEnum Type, long lSize, _variant_t & vtValue)=NULL;
	//删除参数
	virtual void __cdecl ClearAllParameters()=NULL;
	//获得参数
	virtual void __cdecl GetParameterValue(LPCTSTR pszParamName, _variant_t & vtValue)=NULL;
	//获取返回数值
	virtual long __cdecl GetReturnValue()=NULL;
	//执行语句
	virtual bool __cdecl Execute(LPCTSTR pszCommand,bool bRecordset)=NULL;
	//执行命令
	virtual bool __cdecl ExecuteCommand(bool bRecordset)=NULL;
};

//////////////////////////////////////////////////////////////////////////

//ADO 错误类
class CADOError:IADOError
{
	//变量定义
protected:
	enADOErrorType					m_enErrorType;						//错误代号
	TCHAR							m_strErrorDescribe[50];					//错误信息

	//函数定义
public:
	//构造函数
	CADOError();
	//析构函数
	virtual ~CADOError();

	//基础接口
public:
	//释放对象
	virtual bool __cdecl Release() { return true; }
	//是否有效
	virtual bool __cdecl IsValid() { return AfxIsValidAddress(this,sizeof(CADOError))?true:false; }
	//接口查询
	virtual void * __cdecl QueryInterface(const IID & Guid, DWORD dwQueryVer);

	//功能接口
public:
	//错误类型
	virtual enADOErrorType __cdecl GetErrorType() { return m_enErrorType; }
	//错误描述
	virtual LPCTSTR __cdecl GetErrorDescribe() { return m_strErrorDescribe; }
	
	//功能函数
public:
	//设置错误
	void SetErrorInfo(enADOErrorType enErrorType, LPCTSTR pszDescribe);
};


//////////////////////////////////////////////////////////////////////////

//数据库对象
class CXDB : public IXDB
{
	//信息变量
protected:
	CADOError						m_ADOError;							//错误对象
	TCHAR							m_strConnect[100];						//连接字符串
	TCHAR							m_strErrorDescribe[100];					//错误信息

	//状态变量
protected:
	DWORD							m_dwConnectCount;					//重试次数
	DWORD							m_dwConnectErrorTime;				//错误时间
	const DWORD						m_dwResumeConnectCount;				//恢复次数
	const DWORD						m_dwResumeConnectTime;				//恢复时间

	//内核变量
protected:
	_CommandPtr						m_DBCommand;						//命令对象
	_RecordsetPtr					m_DBRecordset;						//记录集对象
	_ConnectionPtr					m_DBConnection;						//数据库对象

	//函数定义
public:
	//构造函数
	CXDB();
	//析构函数
	virtual ~CXDB();

	//基础接口
public:
	//释放对象
	virtual bool __cdecl Release() { if (IsValid()) delete this; return true; }
	//是否有效
	virtual bool __cdecl IsValid() { return AfxIsValidAddress(this,sizeof(CXDB))?true:false; }
	//接口查询
	virtual void * __cdecl QueryInterface(const IID & Guid, DWORD dwQueryVer);

	//管理接口
public:
	//打开连接
	virtual bool __cdecl OpenConnection();
	//关闭记录
	virtual bool __cdecl CloseRecordset();
	//关闭连接
	virtual bool __cdecl CloseConnection();
	//重新连接
	virtual bool __cdecl TryConnectAgain(bool bFocusConnect, CComError * pComError);
	//设置信息
	virtual bool __cdecl SetConnectionInfo(DataBaseType type,LPCTSTR szIP, WORD wPort, LPCTSTR szData, LPCTSTR szName, LPCTSTR szPass);

	//状态接口
public:
	//是否连接错误
	virtual bool __cdecl IsConnectError();
	//是否打开
	virtual bool __cdecl IsRecordsetOpened();

	//记录集接口
public:
	//往下移动
	virtual void __cdecl MoveToNext();
	//移到开头
	virtual void __cdecl MoveToFirst();
	//是否结束
	virtual bool __cdecl IsEndRecordset();
	//获取数目
	virtual long __cdecl GetRecordCount();
	//获取大小
	virtual long __cdecl GetActualSize(LPCTSTR pszParamName);
	//绑定对象
	virtual bool __cdecl BindToRecordset(CADORecordBinding * pBind);

	//字段接口
public:
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, BYTE & bValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, WORD & wValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, INT & nValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, LONG & lValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DWORD & ulValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, UINT & ulValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, DOUBLE & dbValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, __int64 & llValue);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, TCHAR szBuffer[], UINT uSize);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, COleDateTime & Time);
	//获取参数
	virtual bool __cdecl GetFieldValue(LPCTSTR lpFieldName, bool & bValue);

	//命令对象接口
public:
	//设置存储过程
	virtual void __cdecl SetSPName(LPCTSTR pszSpName);
	//插入参数
	virtual void __cdecl AddParamter(LPCTSTR pszName, ADOCG::ParameterDirectionEnum Direction, ADOCG::DataTypeEnum Type, long lSize, _variant_t & vtValue);
	//删除参数
	virtual void __cdecl ClearAllParameters();
	//获得参数
	virtual void __cdecl GetParameterValue(LPCTSTR pszParamName, _variant_t & vtValue);
	//获取返回数值
	virtual long __cdecl GetReturnValue();

	//执行接口
public:
	//执行语句
	virtual bool __cdecl Execute(LPCTSTR pszCommand,bool bRecordset);
	//执行命令
	virtual bool __cdecl ExecuteCommand(bool bRecordset);

	//内部函数
private:
	//获取错误
	LPCTSTR GetComErrorDescribe(CComError & ComError);
	//设置错误
	void SetErrorInfo(enADOErrorType enErrorType, LPCTSTR pszDescribe);
};

//////////////////////////////////////////////////////////////////////////

//数据库引擎辅助类
class CXDBHelper : public CTempldateHelper<IXDB>
{
	//函数定义
public:
	//构造函数
	CXDBHelper(void) : CTempldateHelper<IXDB>(IID_IXDB,
		VER_IXDB,XDB_DLL_NAME,TEXT("CreateXDB")) { }
};

//////////////////////////////////////////////////////////////////////////

#endif