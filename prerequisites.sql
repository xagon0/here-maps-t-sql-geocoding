--enable Ole Automation Procedures / Show Advanced Options
sp_configure 'show advanced options', 1;  
GO  
RECONFIGURE;  
GO  
sp_configure 'Ole Automation Procedures', 1;  
GO  
RECONFIGURE;  
GO  


--disable Ole Automation Procedures / Show Advanced Options
sp_configure 'Ole Automation Procedures', 0;  
GO  
RECONFIGURE;  
GO  
sp_configure 'show advanced options', 0;  
GO  
RECONFIGURE;  
GO  
