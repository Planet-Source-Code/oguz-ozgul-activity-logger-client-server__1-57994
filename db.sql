--BATCH_COUNT=41

--Creating Application Database
CREATE DATABASE dbActivities
GO

--Switching to Application Database
USE dbActivities
GO

--Creating Users table
CREATE TABLE tblUsers(
	intUserID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	strUserName		varchar(64)	NOT NULL,
	strUserPassword		varbinary(64)	NOT NULL,
	strUserFirstName	nvarchar(64)	NOT NULL,
	strUserLastName		nvarchar(64)	NOT NULL,
	strUserTitle		nvarchar(128)	NOT NULL,
	strUserPhone		varchar(32)	NOT NULL,
	strUserExtension	varchar(8)	NOT NULL,
	strUserCellPhone	varchar(32)	NOT NULL,
	strUserEMail		varchar(128)	NOT NULL,
	bitUserIsActive		bit		NOT NULL,
	CONSTRAINT uq_userName UNIQUE NONCLUSTERED (strUserName)
)
GO

--Creating Stored Procedures for user operations
CREATE PROCEDURE prAddUser
	@strUserName		varchar(64),
	@strUserPassword	varchar(64),
	@strUserFirstName	nvarchar(64),
	@strUserLastName	nvarchar(64),
	@strUserTitle		nvarchar(128),
	@strUserPhone		varchar(32),
	@strUserExtension	varchar(8),
	@strUserCellPhone	varchar(32),
	@strUserEMail		varchar(128),
	@intUserID		int OUTPUT,
	@strException		varchar(256) OUTPUT
AS
	IF EXISTS(SELECT intUserID FROM tblUsers (NOLOCK) WHERE strUserName = @strUserName) BEGIN
		SELECT @intUserID = NULL
		SELECT @strException = 'A User with the same user name exists. Please specify another user name'
		RETURN -1
	END
	
	INSERT INTO tblUsers 
		(strUserName, strUserPassword, strUserFirstName, strUserLastName, strUserTitle, strUserPhone, strUserExtension, strUserCellPhone, strUserEMail, bitUserIsActive)
		VALUES
		(@strUserName, CONVERT(VARBINARY(64), pwdEncrypt(@strUserPassword)), @strUserFirstName, @strUserLastName, @strUserTitle, @strUserPhone, @strUserExtension, @strUserCellPhone, @strUserEMail, 1)

	IF @@ERROR = 0 BEGIN
		SELECT @intUserID = @@IDENTITY
		SELECT @strException = NULL
		RETURN 0
	END
	ELSE BEGIN
		SELECT @intUserID = NULL
		SELECT @strException = 'An Exception occured whiile inserting record. The record could not be inserted'
		RETURN -1
	END
GO		

EXEC prAddUser 'administrator', '', 'N/A', 'N/A', 'Activity Organizer Administrator', 'N/A', 'N/A', 'N/A', 'N/A', 0, ''
GO

CREATE PROCEDURE prUpdateUser
	@intUserID		int,
	@strUserFirstName	nvarchar(64),
	@strUserLastName	nvarchar(64),
	@strUserTitle		nvarchar(128),
	@strUserPhone		varchar(32),
	@strUserExtension	varchar(8),
	@strUserCellPhone	varchar(32),
	@strUserEMail		varchar(128),
	@strException		varchar(256) OUTPUT
AS
	DECLARE @strUserName	varchar(64)

	IF NOT EXISTS(SELECT intUserID FROM tblUsers (NOLOCK) WHERE intUserID = @intUserID) BEGIN
		RAISERROR ('User not found', 16, 1)
		SELECT @strException = 'User Not Found'
		RETURN -1
	END

	SELECT @strUserName = strUserName FROM tblUsers (NOLOCK) WHERE intUserID = @intUserID
	IF ISNULL(@strUserName, '') = 'administrator' BEGIN
		RAISERROR('The administrator account can not be updated', 16, 1)
		SELECT @strException = 'The administrator account can not be updated'
		RETURN -1
	END

	UPDATE tblUsers SET
		strUserFirstName = @strUserFirstName,
		strUserLastName = @strUserLastName,
		strUserTitle = @strUserTitle,
		strUserPhone = @strUserPhone,
		strUserExtension = @strUserExtension,
		strUserCellPhone = @strUserCellPhone,
		strUserEMail = @strUserEMail
	WHERE
		intUserID = @intUserID

	IF @@ERROR = 0 BEGIN
		SELECT @strException = NULL
		RETURN 0
	END
	ELSE BEGIN
		SELECT @strException = 'An Exception occured whiile updating record. The record could not be updated'
		RETURN -1
	END
GO		

CREATE PROCEDURE prGetUsers
AS
	SELECT * FROM tblUsers (NOLOCK) WHERE strUserName <> 'administrator' AND bitUserIsActive = 1
GO


CREATE PROCEDURE prGetUser
	@intUserID	int
AS

	SELECT * FROM tblUsers (NOLOCK) WHERE intUserID = @intUserID
GO

CREATE PROCEDURE prRemoveUser
	@intUserID	int
AS
	DECLARE @strUserName	varchar(64)
	SELECT @strUserName = strUserName FROM tblUsers (NOLOCK) WHERE intUserID = @intUserID
	IF ISNULL(@strUserName, '') = 'administrator' BEGIN
		RAISERROR('The administrator account can not be removed', 16, 1)
		RETURN -1
	END

	UPDATE tblUsers SET bitUserIsActive = 0 WHERE intUserID = @intUserID
GO

CREATE PROCEDURE prCheckPassword
	@strUserName		varchar(64),
	@strPassword		varchar(64),
	@intUserID		int OUTPUT
AS
	SELECT @intUserID = intUserID FROM tblUsers (NOLOCK)
	WHERE strUserName = @strUserName AND bitUserIsActive = 1
	--AND pwdCompare(@strPassword, strUserPassword) = 1
GO

--Creating Projects table
CREATE TABLE tblProjects(
	intProjectID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	strProjectName		nvarchar(64)	NOT NULL,
	bitVisible		bit		NOT NULL
)
GO

--Creating Stored Procedures for Project operations
CREATE PROCEDURE prGetProjects
AS
	SELECT intProjectID, strProjectName, bitVisible FROM tblProjects (NOLOCK) ORDER BY strProjectName ASC
GO

CREATE PROCEDURE prAddProject
	@strProjectName		nvarchar(64),
	@intProjectID		int OUTPUT
AS
	INSERT INTO tblProjects (strProjectName, bitVisible) VALUES (@strProjectName, 1)
	IF @@ERROR = 0
		SELECT @intProjectID = @@IDENTITY
	ELSE
		SELECT @intProjectID = NULL
GO	

CREATE PROCEDURE prRemoveProject
	@intProjectID		int
AS
	UPDATE tblProjects SET bitVisible = 0 WHERE intProjectID = @intProjectID
	IF @@ROWCOUNT = 0
		RAISERROR('The project could not be found. Removal aborted', 16, 1)
GO

--Inserting default project records
INSERT INTO tblProjects (strProjectName, bitVisible) VALUES ('No Project', 0)
GO

--Creating Activity Type table
CREATE TABLE tblActivityTypes(
	intActivityTypeID	int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	strActivityTypeName	nvarchar(64)	NOT NULL,
	bitVisible		bit		NOT NULL
)
GO

--Creating Stored Procedures for Activity Type operations
CREATE PROCEDURE prGetActivityTypes
AS
	SELECT intActivityTypeID, strActivityTypeName, bitVisible FROM tblActivityTypes (NOLOCK)
GO

CREATE PROCEDURE prAddActivityType
	@strActivityTypeName	nvarchar(64),
	@intActivityTypeID	int OUTPUT
AS
	INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES (@strActivityTypeName, 1)
	IF @@ERROR = 0
		SELECT @intActivityTypeID = @@IDENTITY
	ELSE
		SELECT @intActivityTypeID = NULL
GO

CREATE PROCEDURE prRemoveActivityType
	@intActivityTypeID	int
AS
	UPDATE tblActivityTypes SET bitVisible = 0 WHERE intActivityTypeID = @intActivityTypeID
	IF @@ROWCOUNT = 0
		RAISERROR('The activity type could not be found. Removal aborted', 16, 1)
GO

--Inserting default Activity Type records
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('Administrative', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('Sales', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('Marketing', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Analysis', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Design', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Development', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Testing', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Deployment', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Documentation', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Maintenance', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Project Management', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Project Meeting', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Migration&Integration', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('PS/Bug Fix', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('Approved Absence', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('LO/CodeWalk', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('LO/Project Training', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('LO/Exam', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('LO/Seminar', 1)
INSERT INTO tblActivityTypes (strActivityTypeName, bitVisible) VALUES ('LO/Training', 1)
GO

--Creating Pause Causes table
CREATE TABLE tblPauseCauses(
	intPauseCauseID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	strPauseCause		nvarchar(128)	NOT NULL,
	bitVisible		bit		NOT NULL
)
GO

--Creating Stored Procedures for Pause Cause operations
CREATE PROCEDURE prAddPauseCause
	@strPauseCause		nvarchar(128),
	@intPauseCauseID	int OUTPUT
AS
	INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES (@strPauseCause, 1)
	IF @@ERROR = 0
		SELECT @intPauseCauseID = @@IDENTITY
	ELSE
		SELECT @intPauseCauseID = NULL
GO

CREATE PROCEDURE prGetPauseCauses
AS
	SELECT intPauseCauseID, strPauseCause, bitVisible FROM tblPauseCauses (NOLOCK)
GO

CREATE PROCEDURE prRemovePauseCause
	@intPauseCauseID	int
AS
	UPDATE tblPauseCauses SET bitVisible = 0 WHERE intPauseCauseID = @intPauseCauseID
	IF @@ROWCOUNT = 0
		RAISERROR('The pause cause could not be found. Removal aborted', 16, 1)
GO

--Inserting default Pause Cause records
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Break - Cigarette', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Break - Tea / Coffee / Snack', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Break - Team Chat', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Break - Lunch - Dinner (During Over-Time)', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Break - Internet Surf', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Break - Other (Please give Detail)', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Phone - Customer', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Phone - Team Lead', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Phone - Project Manager', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Phone - Company', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Phone - Private', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Phone - Other', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Send/Receive', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Reading', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Answering Customer', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Answering Team Lead', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Answering Project Manager', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Answering Company', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Private', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('EMail - Other', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Question/Chat - Customer', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Question/Chat - Team Mate', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Question/Chat - Team Lead', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Question/Chat - Project Manager', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Question/Chat - Other', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Messenger Application Instant Message', 1)
INSERT INTO tblPauseCauses (strPauseCause, bitVisible) VALUES ('Cell Phone', 1)
GO

--Creating Activities table
CREATE TABLE tblActivities(
	intActivityID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	intUserID		int		NOT NULL	REFERENCES tblUsers(intUserID),
	intProjectID		int		NOT NULL	REFERENCES tblProjects(intProjectID),
	intActivityTypeID	int		NOT NULL	REFERENCES tblActivityTypes(intActivityTypeID),
	strActivity		nvarchar(256)	NOT NULL,
	dtmActivityStartDate	datetime	NOT NULL,
	intActivityTotalTimeSec	int		NOT NULL
)
GO

--Creating Stored Procedures for Activity operations
CREATE PROCEDURE prInsertActivity
	@intUserID		int,
	@intProjectID		int,
	@intActivityTypeID	int,
	@strActivity		nvarchar(256),
	@intActivityID		int OUTPUT
AS

	INSERT INTO tblActivities
		(intUserID, intProjectID, intActivityTypeID, strActivity, dtmActivityStartDate, intActivityTotalTimeSec)
		VALUES
		(@intUserID, @intProjectID, @intActivityTypeID, @strActivity, GetDate(), 0)

	IF @@ERROR = 0
		SELECT @intActivityID = @@IDENTITY
	ELSE
		SELECT @intActivityID = NULL
GO

CREATE PROCEDURE prUpdateActivity
	@intActivityID			int,
	@intActivityTotalTimeSec	int
AS
	UPDATE tblActivities SET intActivityTotalTimeSec = @intActivityTotalTimeSec WHERE intActivityID = @intActivityID
	IF @@ROWCOUNT = 0
		RAISERROR('The activity could not be found', 16, 1)
GO

--Creating Pauses table
CREATE TABLE tblPauses(
	intPauseID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	intUserID		int		NOT NULL	REFERENCES tblUsers(intUserID),
	intActivityID		int		NOT NULL	REFERENCES tblActivities(intActivityID),
	intPauseCauseID		int		NOT NULL	REFERENCES tblPauseCauses(intPauseCauseID),
	strPauseCauseDetail	nvarchar(256)	NOT NULL,
	dtmPauseDate		datetime	NOT NULL,
	intPauseTotalTimeSec	int
)
GO	

--Creating Stored Procedures for Pause operations
CREATE PROCEDURE prInsertPause
	@intUserID		int,
	@intActivityID		int,
	@intPauseCauseID	int,
	@strPauseCauseDetail	nvarchar(256),
	@intPauseID		int OUTPUT
AS
	INSERT INTO tblPauses
		(intUserID, intActivityID, intPauseCauseID, strPauseCauseDetail, dtmPauseDate, intPauseTotalTimeSec)
		VALUES
		(@intUserID, @intActivityID, @intPauseCauseID, @strPauseCauseDetail, GetDate(), 0)
	IF @@ERROR = 0
		SELECT @intPauseID = @@IDENTITY
	ELSE
		SELECT @intPauseID = NULL
GO

CREATE PROCEDURE prUpdatePause
	@intPauseID		int,
	@intPauseTotalTimeSec	int
AS
	UPDATE tblPauses SET intPauseTotalTimeSec = @intPauseTotalTimeSec WHERE intPauseID = @intPauseID
	IF @@ROWCOUNT = 0
		RAISERROR('The pause could not be found', 16, 1)
GO

--Creating Logins Log table
CREATE TABLE tblLogins(
	intLoginID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	strUserName		varchar(64)	NOT NULL,
	bitUserNameExists	bit		NOT NULL,
	bitPassCorrect		bit		NOT NULL,
	strRemoteIP		varchar(19)	NOT NULL,
	dtmLoginDate		datetime	NOT NULL
)
GO

--Creating Stored Procedures for Login operations
CREATE PROCEDURE prInsertLogin
	@strUserName		varchar(64),
	@bitPassCorrect		bit,
	@strRemoteIP		varchar(19),
	@intLoginID		int OUTPUT
AS	
	DECLARE @bitUserNameExists	bit

	IF EXISTS(SELECT intUserID FROM tblUsers (NOLOCK) WHERE strUserName = @strUserName)
		SELECT @bitUserNameExists = 1
	ELSE	
		SELECT @bitUserNameExists = 0

	INSERT INTO tblLogins
		(strUserName, bitUserNameExists, bitPassCorrect, strRemoteIP, dtmLoginDate)
		VALUES
		(@strUserName, @bitUserNameExists, @bitPassCorrect, @strRemoteIP, GetDate())

	SELECT @intLoginID = @@IDENTITY
	RETURN 0
GO

--Creating Logouts Log table
CREATE TABLE tblLogOuts(
	intLogOutID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	intUserID		int		NOT NULL	REFERENCES tblUsers(intUserID),
	intLoginID		int		NOT NULL	REFERENCES tblLogins(intLoginID),
	dtmLogOutDate		datetime	NOT NULL,
	bitNormalLogOut		bit		NOT NULL
)
GO

--Creating Stored Procedures For Logout operations
CREATE PROCEDURE prInsertLogOut
	@intUserID		int,
	@intLoginID		int,
	@bitNormalLogOut	bit
AS
	INSERT INTO tblLogOuts
		(intUserID, intLoginID, bitNormalLogOut, dtmLogOutDate)
		VALUES
		(@intUserID, @intLoginID, @bitNormalLogOut, GetDate())
GO

--Creating Exception Log table
CREATE TABLE tblExceptions(
	intExceptionID		int		IDENTITY(1,1)	PRIMARY KEY CLUSTERED,
	strProcedure		varchar(256)	NOT NULL,		
	intExceptionCode	int		NOT NULL,
	strExceptionMessage	nvarchar(512)	NOT NULL,
	strExtraInfo		nvarchar(512)	NOT NULL,
	bitIsRunTimeException	bit		NOT NULL,
	dtmExceptionDate	datetime	NOT NULL
)
GO

--Creating Stored Procedures for Exception Logging operations
CREATE PROCEDURE prLogException
	@strProcedure		varchar(256),
	@intExceptionCode	int,
	@strExceptionMessage	nvarchar(512),
	@strExtraInfo		nvarchar(512),
	@bitIsRunTimeException	bit
AS

	INSERT INTO tblExceptions
		(strProcedure, intExceptionCode, strExceptionMessage, strExtraInfo, bitIsRunTimeException, dtmExceptionDate)
		VALUES
		(@strProcedure, @intExceptionCode, @strExceptionMessage, @strExtraInfo, @bitIsRunTimeException, Getdate())
GO

--Creating Stored Procedures for Reporting operations
CREATE PROCEDURE prReportsGetUserDailyActivityReport
	@intUserID	int,
	@dtmDayStart	char(17),
	@dtmDayEnd	char(17)
AS
	--SELECT intUserID, intProjectID, intActivityTypeID, strActivity, dtmActivityStartDate, intActivityTotalTimeSec
	SELECT strActivity, dtmActivityStartDate, intActivityTotalTimeSec
	FROM tblActivities (NOLOCK)
	WHERE intUserID = @intUserID AND dtmActivityStartDate BETWEEN @dtmDayStart AND @dtmDayEnd
	ORDER BY dtmActivityStartDate ASC
GO

CREATE PROCEDURE prReportsGetAllDailyActivityReports
	@dtmDayStart	char(8),
	@dtmDayEnd	char(8)
AS
	--SELECT intUserID, intProjectID, intActivityTypeID, strActivity, dtmActivityStartDate, intActivityTotalTimeSec
	SELECT intUserID, strActivity, dtmActivityStartDate, intActivityTotalTimeSec
	FROM tblActivities (NOLOCK)
	WHERE dtmActivityStartDate BETWEEN @dtmDayStart AND @dtmDayEnd
	ORDER BY intUserID ASC, dtmActivityStartDate ASC
GO

CREATE PROCEDURE prReportsGetUserDailyPauseReport
	@intUserID	int,
	@dtmDay		char(8)
AS
	SELECT intUserID, intActivityID, intPauseCauseID, strPauseCauseDetail, dtmPauseDate, intPauseTotalTimeSec
	FROM tblPauses (NOLOCK)
	WHERE intUserID = @intUserID AND DATEDIFF(d, dtmPauseDate, @dtmDay) = 0
	ORDER BY dtmPauseDate ASC
GO

CREATE PROCEDURE prReportsGetAllDailyPauseReports
	@dtmDay		char(8)
AS
	SELECT intUserID, intActivityID, intPauseCauseID, strPauseCauseDetail, dtmPauseDate, intPauseTotalTimeSec
	FROM tblPauses (NOLOCK)
	WHERE DATEDIFF(d, dtmPauseDate, @dtmDay) = 0
	ORDER BY intUserID ASC, dtmPauseDate ASC
GO