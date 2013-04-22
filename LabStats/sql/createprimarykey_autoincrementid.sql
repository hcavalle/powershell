alter table general.dbo.Lab_Logon
add id int identity(1,1)

/*second part*/
alter table general.dbo.Lab_Logon add constraint PK_id
PRIMARY KEY clustered (id)