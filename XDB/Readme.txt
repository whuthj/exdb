xdb v1.0 for VC
����:
	1.��װ�� VC ���ݿ� sql ��䡢�洢���̲���;
	2.����õ����ݿⷵ��ֵ;
	3.֧�� Sqlserver��Oracle,��������������֧�ָ���;
	4.VC ����ӿ����;
����:����
����:whuthj@163.com;
Code license:GNU Lesser GPL;
˵��:
	���� VC �������ݿ���ԱȽ��鷳,�� xdb �������ݿ�ǳ�����,
	��Դ��������⴫�������޸�����,ּ�ڷ��㿪����ʹ��.
����ʾ��:
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