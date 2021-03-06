/****** Object:  Table [dbo].[tblCategory]    Script Date: 2/16/2003 6:52:12 PM ******/
CREATE TABLE [dbo].[tblCategory] (
	[iCategoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[vCategoryName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblLogin]    Script Date: 2/16/2003 6:52:15 PM ******/
CREATE TABLE [dbo].[tblLogin] (
	[vLoginID] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vPassword] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vDesignation] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vStatus] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblMarketing]    Script Date: 2/16/2003 6:52:15 PM ******/
CREATE TABLE [dbo].[tblMarketing] (
	[iAutoID] [int] IDENTITY (1, 1) NOT NULL ,
	[dtKnockingDate] [datetime] NULL ,
	[iCategoryID] [int] NULL ,
	[vCompanyName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vAddress] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vContactPerson] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vKnockingBy] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iResponsePercent] [int] NULL ,
	[vSoftwareType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vResponse] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vDemoBy] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtDemoDate] [datetime] NULL ,
	[vDemoResponse] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vSubmitProposal] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[mProposalPrice] [money] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblCategory] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblCategory] PRIMARY KEY  CLUSTERED 
	(
		[iCategoryID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLogin] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblLogin] PRIMARY KEY  CLUSTERED 
	(
		[vLoginID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMarketing] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblMarketing] PRIMARY KEY  CLUSTERED 
	(
		[vCompanyName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMarketing] ADD 
	CONSTRAINT [FK_tblMarketing_tblCategory] FOREIGN KEY 
	(
		[iCategoryID]
	) REFERENCES [dbo].[tblCategory] (
		[iCategoryID]
	) ON UPDATE CASCADE ,
	CONSTRAINT [FK_tblMarketing_tblLogin] FOREIGN KEY 
	(
		[vKnockingBy]
	) REFERENCES [dbo].[tblLogin] (
		[vLoginID]
	) ON UPDATE CASCADE 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.SP_Add_tblLogin    Script Date: 2/16/2003 6:52:16 PM ******/
CREATE PROCEDURE SP_Add_tblLogin

	@vLoginID 	varchar(25),
	@vPassword 	varchar(10),
	@vDesignation	varchar(25),
	@vStatus	varchar(10)
	
AS
	INSERT INTO tblLogin (vLoginId,vPassword, vDesignation, vStatus)

	Values(@vLoginId,@vPassword,@vDesignation, @vStatus)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.SP_Add_tblMarketing    Script Date: 2/16/2003 6:52:16 PM ******/
CREATE PROCEDURE SP_Add_tblMarketing

	@iCategoryID 		int,
	@vCompanyName 	varchar(50),
	@vAddress		varchar(100),
	@vPhone		varchar(50),
	@vContactPerson	varchar(50),
	@dtKnockingDate	datetime,
	@vKnockingBy		varchar(25),
	@iResponsePercent	int,
	@vSoftwareType	varchar(25),
	@vResponse		varchar(100),
	@vDemoBy		varchar(50),
	@dtDemoDate		datetime,
	@vDemoResponse	varchar(100),
	@vSubmitProposal	varchar(5),
	@mProposalPrice	money
	
AS
	INSERT INTO tblMarketing (iCategoryID, vCompanyName, vAddress, vPhone, vContactPerson, 
					dtKnockingDate,vKnockingBy, iResponsePercent, vSoftwareType, vResponse,
					vDemoBy, dtDemoDate, vDemoResponse, vSubmitProposal, mProposalPrice)

	Values	(@iCategoryID, @vCompanyName, @vAddress, @vPhone, @vContactPerson,
				@dtKnockingDate, @vKnockingBy, @iResponsePercent, @vSoftwareType, @vResponse,
				@vDemoBy, @dtDemoDate, @vDemoResponse, @vSubmitProposal, @mProposalPrice)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.SP_Up_tblLogin    Script Date: 2/16/2003 6:52:16 PM ******/
CREATE PROCEDURE SP_Up_tblLogin
	
	@vLoginID 	varchar(25),
	@vPassword 	varchar(10),
	@vDesignation	varchar(25),
	@vStatus	varchar(10)

AS

	UPDATE tblLogin
	SET 
		vPassword	=	@vPassword,
		vDesignation	=	@vDesignation,
		vStatus		=	@vStatus

	WHERE vLoginID	=	@vLoginID
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.Sp_Up_tblMarketing    Script Date: 2/16/2003 6:52:16 PM ******/
CREATE  PROCEDURE Sp_Up_tblMarketing

	@iAutoID		int,
	@iCategoryID 		int,
	@vCompanyName 	varchar(50),
	@vAddress		varchar(100),
	@vPhone		varchar(50),
	@vContactPerson	varchar(50),
	@dtKnockingDate	datetime,
	@vKnockingBy		varchar(25),
	@iResponsePercent	int,
	@vSoftwareType	varchar(25),
	@vResponse		varchar(100),
	@vDemoBy		varchar(50),
	@dtDemoDate		datetime,
	@vDemoResponse	varchar(100),
	@vSubmitProposal	varchar(5),
	@mProposalPrice	money

AS

	Update tblMarketing

	Set	iCategoryID		=	@iCategoryID,
		vCompanyName		=	@vCompanyName,
		vAddress		=	@vAddress,
		vPhone			=	@vPhone,
		vContactPerson		=	@vContactPerson,
		dtKnockingDate		=	@dtKnockingDate,
		vKnockingBy		=	@vKnockingBy,
		iResponsePercent	=	@iResponsePercent,
		vSoftwareType		=	@vSoftwareType,
		vResponse		=	@vResponse,
		vDemoBy		=	@vDemoBy,
		dtDemoDate		=	@dtDemoDate,
		vDemoResponse	=	@vDemoResponse,
		vSubmitProposal		=	@vSubmitProposal,
		mProposalPrice		=	@mProposalPrice

	Where	iAutoID			=	@iAutoID
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

