/****** Object:  Table [dbo].[tblDepartment]    Script Date: 11/28/2002 7:09:06 PM ******/
CREATE TABLE [dbo].[tblDepartment] (
	[siDeptID] [smallint] IDENTITY (1, 1) NOT NULL ,
	[vDeptName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[siGroupID] [smallint] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblGroup]    Script Date: 11/28/2002 7:09:07 PM ******/
CREATE TABLE [dbo].[tblGroup] (
	[siGroupID] [smallint] IDENTITY (1, 1) NOT NULL ,
	[vGroupName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblDepartment] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblDepartment] PRIMARY KEY  CLUSTERED 
	(
		[siDeptID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblGroup] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblGroup] PRIMARY KEY  CLUSTERED 
	(
		[siGroupID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblDepartment] ADD 
	CONSTRAINT [FK_tblDepartment_tblGroup] FOREIGN KEY 
	(
		[siGroupID]
	) REFERENCES [dbo].[tblGroup] (
		[siGroupID]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblDept    Script Date: 11/28/2002 7:09:07 PM ******/
CREATE  PROCEDURE Sp_Add_tblDept
		
		@vDeptName	varchar(25),
		@siGroupID	smallint
				
AS


	INSERT INTO tblDepartment (vDeptName, siGroupID)

	Values (@vDeptName, @siGroupID)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblGroup    Script Date: 11/28/2002 7:09:07 PM ******/
CREATE  PROCEDURE Sp_Add_tblGroup
		
		@vGroupName	varchar(25)
				
AS


	INSERT INTO tblGroup (vGroupName)

	Values (@vGroupName)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblDept    Script Date: 11/28/2002 7:09:07 PM ******/
CREATE PROCEDURE Sp_Up_tblDept

	@siDeptID	smallint,
	@vDeptName	varchar(25),
	@siGroupID	smallint

AS

	UPDATE tblDepartment
	SET 
		vDeptName	=	@vDeptName,
		siGroupID	=	@siGroupID

	WHERE siDeptID	= 	@siDeptID
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblGroup    Script Date: 11/28/2002 7:09:07 PM ******/
CREATE PROCEDURE Sp_Up_tblGroup

	@siGroupID	smallint,
	@vGroupName	varchar(25)

AS
	UPDATE tblGroup
	SET 
		vGroupName        =         @vGroupName

	WHERE siGroupID	 =	@siGroupID
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

