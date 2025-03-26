USE SRLSQL;

--Queries for grant permissions to juliandevuser 

--Command to initialize the local pipe connection 
-- np:\\.\pipe\LOCALDB#6D074CBC\tsql\query

SELECT 
    perm.state_desc AS Estado,
    perm.permission_name AS Permiso,
    obj.name AS Objeto
FROM sys.database_permissions perm
JOIN sys.objects obj ON perm.major_id = obj.object_id
JOIN sys.database_principals prin ON perm.grantee_principal_id = prin.principal_id
WHERE prin.name = 'juliandevuser';
GO


ALTER ROLE db_datareader ADD MEMBER juliandevuser;  -- Permiso de lectura
ALTER ROLE db_datawriter ADD MEMBER juliandevuser;  -- Permiso de escritura
GO

GRANT CREATE TABLE TO juliandevuser;  -- Permiso para crear tablas
GRANT EXECUTE TO juliandevuser;      -- Permiso para ejecutar procedimientos almacenados
GO

