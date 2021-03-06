/****** Object:  Table [dbo].[CBtblEmpDetails]    Script Date: 1/15/2003 3:19:54 PM ******/
CREATE TABLE [dbo].[CBtblEmpDetails] (
	[vEmpId] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vEmpName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vFName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPresentAdd] [varchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPmtAdd] [varchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iStatus] [int] NULL ,
	[vPhone] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vEmail] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtBD] [datetime] NULL ,
	[dtHD] [datetime] NULL ,
	[iDeptId] [int] NOT NULL ,
	[vDesignation] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[mSalary] [money] NULL ,
	[iPhoto] [image] NULL ,
	[vPassword] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vLogIn] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vLogOut] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

/****** Object:  Table [dbo].[CBtblEmpDetails1]    Script Date: 1/15/2003 3:19:54 PM ******/
CREATE TABLE [dbo].[CBtblEmpDetails1] (
	[vEmpId] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vEmpName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vFName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPresentAdd] [varchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPmtAdd] [varchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[siStatus] [smallint] NULL ,
	[vPhone] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vEmail] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtBD] [datetime] NULL ,
	[dtHD] [datetime] NULL ,
	[siDeptId] [smallint] NOT NULL ,
	[vDesignation] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[mSalary] [money] NULL ,
	[iPhoto] [image] NULL ,
	[vPassword] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vLogIn] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vLogOut] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vActive] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

/****** Object:  Table [dbo].[CBtblLogin]    Script Date: 1/15/2003 3:19:54 PM ******/
CREATE TABLE [dbo].[CBtblLogin] (
	[vEmpId] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vPassword] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[nRole] [smallint] NULL ,
	[vCreator] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtDate] [datetime] NULL ,
	[vActive] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[siCrime] [smallint] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblAttend]    Script Date: 1/15/2003 3:19:55 PM ******/
CREATE TABLE [dbo].[tblAttend] (
	[iID] [int] IDENTITY (1, 1) NOT NULL ,
	[vEmpID] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtDate] [datetime] NULL ,
	[dtLogIn] [datetime] NULL ,
	[vActualLogIn] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtLogOut] [datetime] NULL ,
	[vActualLogOut] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iOverTime] [real] NULL ,
	[vRemark] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vStatus] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vUpdateBy] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblCompanyInfo]    Script Date: 1/15/2003 3:19:55 PM ******/
CREATE TABLE [dbo].[tblCompanyInfo] (
	[vCompanyName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vAddress] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vFax] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vEmail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vWebSite] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vEmpID] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblDepartment]    Script Date: 1/15/2003 3:19:55 PM ******/
CREATE TABLE [dbo].[tblDepartment] (
	[siDeptID] [smallint] IDENTITY (1, 1) NOT NULL ,
	[vDeptName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[siGroupID] [smallint] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblEmpDetails]    Script Date: 1/15/2003 3:19:55 PM ******/
CREATE TABLE [dbo].[tblEmpDetails] (
	[vEmpId] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vEmpName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vFName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPresentAdd] [varchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPmtAdd] [varchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPhone] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vEmail] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtBD] [datetime] NULL ,
	[dtHD] [datetime] NULL ,
	[siDeptID] [smallint] NULL ,
	[vDesignation] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[siStatus] [smallint] NULL ,
	[mSalary] [money] NULL ,
	[iPhoto] [image] NULL ,
	[vLogIn] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vLogOut] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPassword] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vActive] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblGroup]    Script Date: 1/15/2003 3:19:55 PM ******/
CREATE TABLE [dbo].[tblGroup] (
	[siGroupID] [smallint] IDENTITY (1, 1) NOT NULL ,
	[vGroupName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblLoginAA]    Script Date: 1/15/2003 3:19:56 PM ******/
CREATE TABLE [dbo].[tblLoginAA] (
	[vLoginId] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vPassword] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[siRole] [smallint] NULL ,
	[dtDate] [datetime] NULL ,
	[siPinChk] [smallint] NULL ,
	[vStatus] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[siCrime] [smallint] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblObject]    Script Date: 1/15/2003 3:19:56 PM ******/
CREATE TABLE [dbo].[tblObject] (
	[vObjectName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vObjectCaption] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vDefault] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblObjectPermission]    Script Date: 1/15/2003 3:19:56 PM ******/
CREATE TABLE [dbo].[tblObjectPermission] (
	[iAutoID] [int] IDENTITY (1, 1) NOT NULL ,
	[vEmpID] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vObjectName] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vEnable] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblRemark]    Script Date: 1/15/2003 3:19:56 PM ******/
CREATE TABLE [dbo].[tblRemark] (
	[iAutoID] [int] IDENTITY (1, 1) NOT NULL ,
	[vRemark] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblAttend] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblAttend] PRIMARY KEY  CLUSTERED 
	(
		[iID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblCompanyInfo] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblCompanyInfo] PRIMARY KEY  CLUSTERED 
	(
		[vCompanyName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblDepartment] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblDepartment] PRIMARY KEY  CLUSTERED 
	(
		[siDeptID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblEmpDetails] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblEmpDetails] PRIMARY KEY  CLUSTERED 
	(
		[vEmpId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblGroup] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblGroup] PRIMARY KEY  CLUSTERED 
	(
		[siGroupID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblObject] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblObject] PRIMARY KEY  CLUSTERED 
	(
		[vObjectName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblObjectPermission] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblObjectPermission] PRIMARY KEY  CLUSTERED 
	(
		[iAutoID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblRemark] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblRemark] PRIMARY KEY  CLUSTERED 
	(
		[iAutoID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblAttend] ADD 
	CONSTRAINT [FK_tblAttend_tblEmpDetails] FOREIGN KEY 
	(
		[vEmpID]
	) REFERENCES [dbo].[tblEmpDetails] (
		[vEmpId]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblCompanyInfo] ADD 
	CONSTRAINT [FK_tblCompanyInfo_tblEmpDetails] FOREIGN KEY 
	(
		[vEmpID]
	) REFERENCES [dbo].[tblEmpDetails] (
		[vEmpId]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblDepartment] ADD 
	CONSTRAINT [FK_tblDepartment_tblGroup] FOREIGN KEY 
	(
		[siGroupID]
	) REFERENCES [dbo].[tblGroup] (
		[siGroupID]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblEmpDetails] ADD 
	CONSTRAINT [FK_tblEmpDetails_tblDepartment] FOREIGN KEY 
	(
		[siDeptID]
	) REFERENCES [dbo].[tblDepartment] (
		[siDeptID]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblObjectPermission] ADD 
	CONSTRAINT [FK_tblObjectPermission_tblEmpDetails] FOREIGN KEY 
	(
		[vEmpID]
	) REFERENCES [dbo].[tblEmpDetails] (
		[vEmpId]
	) ON UPDATE CASCADE ,
	CONSTRAINT [FK_tblObjectPermission_tblObject] FOREIGN KEY 
	(
		[vObjectName]
	) REFERENCES [dbo].[tblObject] (
		[vObjectName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.vrDateWiseAttendant    Script Date: 1/15/2003 3:19:56 PM ******/

/****** Object:  View dbo.vrDateWiseAttendant    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE VIEW dbo.vrDateWiseAttendant
AS
SELECT     dbo.tblAttend.vEmpID, dbo.tblAttend.dtDate, dbo.tblAttend.dtLogIn, dbo.tblAttend.dtLogOut, dbo.tblEmpDetails.vEmpName, 
                      dbo.tblEmpDetails.vDesignation
FROM         dbo.tblAttend INNER JOIN
                      dbo.tblEmpDetails ON dbo.tblAttend.vEmpID = dbo.tblEmpDetails.vEmpId


GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblCompanyInfo    Script Date: 1/15/2003 3:19:56 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Add_tblCompanyInfo    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE PROCEDURE Sp_Add_tblCompanyInfo

	@vCompanyName	varchar(50),
	@vAddress		varchar(200),
	@vPhone		varchar(50),
	@vFax			varchar(50),
	@vEmail		varchar(50),
	@vWebSite		varchar(50),
	@vEmpID		varchar(25)

AS

	Insert Into	tblCompanyInfo	(vCompanyName, vAddress, vPhone, vFax, vEmail, vWebSite, vEmpID)

	Values	(@vCompanyName, @vAddress, @vPhone, @vFax, @vEmail, @vWebSite, @vEmpID)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblDept    Script Date: 1/15/2003 3:19:56 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Add_tblDept    Script Date: 1/13/2003 12:22:23 PM ******/
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

/****** Object:  Stored Procedure dbo.Sp_Add_tblEmpDetails    Script Date: 1/15/2003 3:19:56 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Add_tblEmpDetails    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_Add_tblEmpDetails
		
		@vEmpId      	varchar(25), 
		@vEmpName	varchar(25),
		@vDesignation	varchar(25),
		@siDeptId	smallint,
		@vFName	varchar(25),
		@vPresentAdd	varchar(256),
		@vPmtAdd	varchar(256),
		@vPhone	varchar(25),
		@vEmail	varchar(200),
		@dtBD		datetime,
		@dtHD		datetime,
		@siStatus	int,
		@mSalary	money,
		@vLogIn	varchar(25),
		@vLogOut	varchar(25),
		@vPassword	varchar(10),
		@vActive	varchar(10)

--		@iPhoto	image	


AS

	INSERT INTO tblEmpDetails (vEmpID,vEmpName,vDesignation, siDeptID, vFName,vPresentAdd,vPmtAdd,vPhone,vEmail,dtBD,dtHD,siStatus, mSalary, vLogIn, vLogOut, vPassword, vActive)
	Values (@vEmpID, @vEmpName, @vDesignation, @siDeptID, @vFName, @vPresentAdd, @vPmtAdd, @vPhone, @vEmail, @dtBD, @dtHD, @siStatus, @mSalary, @vLogIn, @vLogOut, @vPassword, @vActive)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblGroup    Script Date: 1/15/2003 3:19:56 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Add_tblGroup    Script Date: 1/13/2003 12:22:23 PM ******/
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

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblObject    Script Date: 1/15/2003 3:19:56 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Add_tblObject    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_Add_tblObject

		@vObjectName		varchar(25),
		@vObjectCaption	varchar(50),
		@vDefault		varchar(10)

AS

	INSERT INTO tblObject (vObjectName, vObjectCaption, vDefault)

	Values	(@vObjectName, @vObjectCaption, @vDefault)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblObjectPermission    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Add_tblObjectPermission    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_Add_tblObjectPermission

		@vEmpID	varchar(25),
		@vObjectName	varchar(50),
		@vEnable	varchar(10)

AS

	INSERT INTO tblObjectPermission (vEmpID, vObjectName, vEnable)

	Values	(@vEmpID, @vObjectName, @vEnable)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Add_tblRemark    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Add_tblRemark    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_Add_tblRemark
		
		@vRemark	varchar(100)
				
AS


	INSERT INTO tblRemark (vRemark)

	Values (@vRemark)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_LogIn_tblAttend    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_LogIn_tblAttend    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_LogIn_tblAttend
		
		@vEmpID		varchar(25),
		@vActualLogIn		varchar(25),
		@vActualLogOut	varchar(25),
		@vRemark		varchar(100)
					
AS


	declare @dtDate as varchar(20)
	declare @vGetDate as varchar(20)
	
	select	@vGetDate = getdate()
	select @dtDate=substring(@vGetDate,1,11)

	INSERT INTO tblAttend (vEmpID, dtDate, dtLogIn, vActualLogin, vActualLogOut, iOverTime, vRemark, vStatus)

	Values (@vEmpID, @dtDate, GetDate(), @vActualLogIn, @vActualLogOut, 0, @vRemark, "Running")

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_LogOut_tblAttend    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_LogOut_tblAttend    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_LogOut_tblAttend
		
		@vEmpID		varchar(25)
						
AS


	declare @dtDate as varchar(20)
	declare @vGetDate as varchar(20)
	
	select	@vGetDate = getdate()
	select @dtDate=substring(@vGetDate,1,11)

	Update 	tblAttend

	Set	dtLogOut	=	Getdate()

	Where	vEmpID	=	@vEmpID	And	dtDate	=	@dtDate

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Manual_LogIn_tblAttend    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Manual_LogIn_tblAttend    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_Manual_LogIn_tblAttend
		
		@vEmpID		varchar(25),
		@dtDate		datetime,
		@dtLogin		datetime,
		@vActualLogIn		varchar(25),
		@vActualLogOut	varchar(25),
		@iOverTime		real,
		@vRemark		varchar(100),
		@vStatus		varchar(25),
		@vUpdateBy		varchar(25)
					
AS

	INSERT INTO tblAttend (vEmpID, dtDate, dtLogIn, vActualLogin, vActualLogOut, iOverTime, vRemark, vStatus, vUpdateBy)

	Values (@vEmpID,  @dtDate, @dtLogin, @vActualLogIn, @vActualLogOut, @iOverTime, @vRemark, @vStatus, @vUpdateBy)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblAttend    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Up_tblAttend    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE PROCEDURE Sp_Up_tblAttend

	@iID		int,
	@dtDate	datetime,
	@dtLogIn	datetime,
	@dtLogOut	datetime,
	@iOverTime	real,
	@vRemark	varchar(100),
	@vStatus	varchar(25),
	@vUpdateBy	varchar(25)

AS

	UPDATE tblAttend
	SET 
		dtDate		=	@dtDate,
		dtLogIn		=	@dtLogIn,
		dtLogOut	=	@dtLogOut,
		iOverTime	=	@iOverTime,
		vRemark	=	@vRemark,
		vStatus		=	@vStatus,
		vUpdateBy	=	@vUpdateBy

	WHERE iID		= 	@iID

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblCompanyInfo    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Up_tblCompanyInfo    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE PROCEDURE Sp_Up_tblCompanyInfo

	@vCompanyName	varchar(50),
	@vAddress		varchar(200),
	@vPhone		varchar(50),
	@vFax			varchar(50),
	@vEmail		varchar(50),
	@vWebSite		varchar(50),
	@vEmpId		varchar(25)
AS

	Update		tblCompanyInfo

	Set	vCompanyName		=	@vCompanyName,
		vAddress		=	@vAddress,
		vPhone			=	@vPhone,
		vFax			=	@vFax,
		vEmail			=	@vEmail,
		vWebSite		=	@vWebSite,
		vEmpID			=	@vEmpID

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblDept    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Up_tblDept    Script Date: 1/13/2003 12:22:23 PM ******/
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

/****** Object:  Stored Procedure dbo.Sp_Up_tblEmpDetails    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Up_tblEmpDetails    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE PROCEDURE Sp_Up_tblEmpDetails
	
		@vEmpId      	varchar(25), 
		@vEmpName	varchar(25),
		@vDesignation	varchar(25),
		@siDeptId	smallint,
		@vFName	varchar(25),
		@vPresentAdd	varchar(256),
		@vPmtAdd	varchar(256),
		@vPhone	varchar(25),
		@vEmail	varchar(200),
		@dtBD		datetime,
		@dtHD		datetime,
		@siStatus	smallint,
		@mSalary	money,
		@vLogIn	varchar(25),
		@vLogOut	varchar(25),
		@vPassword	varchar(10),
		@vActive	varchar(10)

AS
	UPDATE tblEmpDetails
	SET 

	      
	vEmpName    	=	@vEmpName,	
	vDesignation	= 	@vDesignation,
	siDeptId 	=	@siDeptId,
	vFName	=	@vFName,
	vPresentAdd   	=	@vPresentAdd,
	vPmtAdd	=	@vPmtAdd,
	vPhone		=	@vPhone,
	vEmail		=	@vEmail,
	dtBD		=	@dtBD,
	dtHD		=	@dtHD,			
	siStatus		=	@siStatus,
	mSalary		=	@mSalary,
	vLogIn		=	@vLogIn,
	vLogOut	=	@vLogOut,
	vPassword	=	@vPassword,
	vActive		=	@vActive

	WHERE vEmpId         =	@vEmpId

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblGroup    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Up_tblGroup    Script Date: 1/13/2003 12:22:23 PM ******/
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

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblObject    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Up_tblObject    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_Up_tblObject

		@vObjectName		varchar(25),
		@vObjectCaption	varchar(50),
		@vDefault		varchar(10)

AS

	Update	tblObject

	Set	vObjectCaption	=	@vObjectCaption,

		vDefault	=	@vDefault

	Where	vObjectName	=	@vObjectName

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblObjectPermission    Script Date: 1/15/2003 3:19:57 PM ******/

/****** Object:  Stored Procedure dbo.Sp_Up_tblObjectPermission    Script Date: 1/13/2003 12:22:23 PM ******/
CREATE  PROCEDURE Sp_Up_tblObjectPermission

		@vEmpID	varchar(25),
		@vObjectName	varchar(50),
		@vEnable	varchar(10)

AS

	INSERT INTO tblObjectPermission (vEmpID, vObjectName, vEnable)

	Values	(@vEmpID, @vObjectName, @vEnable)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

