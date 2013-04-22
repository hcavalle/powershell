create view LabLogonView
as
SELECT ALL 
       [computername]
      ,[username]
      ,[logonid]
      ,[logontime]
      ,[logofftime]
      ,[sessionlength]
      ,[id]
  FROM [Lab_Logon]
go
