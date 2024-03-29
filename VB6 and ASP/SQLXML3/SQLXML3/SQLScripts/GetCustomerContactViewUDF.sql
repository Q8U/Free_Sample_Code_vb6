if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetCustomerContactView]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetCustomerContactView]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE FUNCTION GetCustomerContactView ( @CustomerID char(5) )
RETURNS @CustomerContactView TABLE
   (
    CustomerID     nchar(5),
    ContactName   nvarchar(30)
   )
AS
BEGIN
   INSERT @CustomerContactView
        SELECT CustomerID, ContactName
        FROM Customers 
        WHERE CustomerID = @CustomerID
   RETURN
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

