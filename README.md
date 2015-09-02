xdb v1.0 for VC
特性:
	1.封装了 VC 数据库 sql 语句、存储过程操作;
	2.方便得到数据库返回值;
	3.支持 Sqlserver、Oracle,后续将继续更新支持更多;
	4.VC 面向接口设计;
作者:胡俊
邮箱:whuthj@163.com;
Code license:GNU Lesser GPL;
说明:
	由于 VC 操作数据库相对比较麻烦,而 xdb 操作数据库非常方便,
	本源码可以随意传播或者修改完善,旨在方便开发者使用.
代码示例:


	CXDBHelper m_db;
	if ((m_db.GetInterface()==NULL)&&(m_db.CreateInstance()==false))
	{
		
	}
	DataBaseType type=MSSQLSERVER;
	m_db->SetConnectionInfo(type,"localhost",1433,"DBTest","sa","123456");
	LPCTSTR strSql=TEXT("SELECT * FROM tbTest");
	if(m_db->OpenConnection())
	{
		m_db->Execute(strSql,true);
	}
	
	TCHAR szName[50];
	m_db->GetFieldValue("UserId",szName,sizeof(szName));

	GetDlgItem(IDC_EDIT_DATA)->SetWindowText(szName);
