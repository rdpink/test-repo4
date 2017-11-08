update users
set password = 'QBlD5jzPE/RF42H3vqi/WBVilaocX7bVvg==',
loginid = 'admin',
IsActive = 1,
IsLocked = 0
where userid = 1

update EmailAccounts
set IsActive = 0

update GlobalSettings
set [GlobalSettingsXml].modify('replace value of (/GlobalSettings/SmtpSettings/EmailQueueMaxSendAttempts/text())[1] with ''0''')

update GlobalSettings
set [GlobalSettingsXml].modify('replace value of (/GlobalSettings/SmtpSettings/SmtpServer/text())[1] with ''invalid''')

update GlobalSettings
  set GlobalSettingsOther('test')
set GlobalSettingsXml.modify('
 delete /GlobalSettings/ZendeskSettings[1]')
