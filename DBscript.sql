USE [YourDB]
GO

/****** Object:  Table [dbo].[CalcResultView]    Script Date: 18.03.2023 23:15:52 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CalcResultView](
	[ID] [int] NULL,
	[NParameterTotal] [nvarchar](50) NULL,
	[NStatistically] [nvarchar](50) NULL,
	[PercentStatistically] [nvarchar](50) NULL,
	[DoNotFitStatistically] [nvarchar](50) NULL,
	[CalcID] [nvarchar](50) NULL,
	[User] [nvarchar](50) NULL,
	[TimePointData] [nvarchar](50) NULL,
	[TimePointCalc] [nvarchar](50) NULL,
	[Note] [nvarchar](max) NULL,
	[Active] [bit] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


CREATE TABLE [dbo].[CalcRow](
	[CalcID] [nvarchar](50) NULL,
	[VIRT_OZID] [nvarchar](50) NULL,
	[Total N] [nvarchar](50) NULL,
	[KPI0] [nvarchar](50) NULL,
	[KPI1] [nvarchar](50) NULL,
	[KPI2] [nvarchar](50) NULL,
	[KPI3] [nvarchar](50) NULL,
	[FitStatistically] [nvarchar](50) NULL,
	[RelevantForDiscussion] [nvarchar](50) NULL,
	[Note] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


CREATE TABLE [dbo].[CalcRowSearch](
	[CalcID] [nvarchar](50) NULL,
	[VIRT_OZID] [nvarchar](50) NULL,
	[TotalN0] [nvarchar](50) NULL,
	[TotalN1] [nvarchar](50) NULL,
	[TotalN2] [nvarchar](50) NULL,
	[TotalN3] [nvarchar](50) NULL,
	[KPI0] [nvarchar](50) NULL,
	[KPI1] [nvarchar](50) NULL,
	[KPI2] [nvarchar](50) NULL,
	[KPI3] [nvarchar](50) NULL,
	[FitStatistically] [nvarchar](50) NULL,
	[RelevantForDiscussion] [nvarchar](50) NULL,
	[Active] [nvarchar](50) NULL,
	[Note] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[CalcRowSearch] ADD  DEFAULT ('false') FOR [Active]
GO




CREATE TABLE [dbo].[CalculationRaw](
	[CalcID] [nvarchar](50) NOT NULL,
	[CalcDate] [nvarchar](50) NULL,
	[ProductCode] [nvarchar](50) NULL,
	[TREND_WERT] [nvarchar](50) NULL,
	[TREND_WERT_2] [nvarchar](50) NULL,
	[CL] [nvarchar](50) NULL,
	[LCL] [nvarchar](50) NULL,
	[UCL] [nvarchar](50) NULL,
	[LAL] [nvarchar](50) NULL,
	[UAL] [nvarchar](50) NULL,
	[TS_ABS] [nvarchar](50) NULL,
	[SORT_DATE] [nvarchar](50) NULL,
	[LAUFNR] [nvarchar](50) NULL,
	[EXCURSION] [nvarchar](50) NULL,
	[VIRT_OZID] [nvarchar](50) NULL,
	[VALUE] [nvarchar](50) NULL,
	[BatchID] [nvarchar](50) NULL,
	[lowSD] [nvarchar](50) NULL,
	[uppSD] [nvarchar](50) NULL,
	[mu] [nvarchar](50) NULL,
	[sigma] [nvarchar](50) NULL,
	[upp] [nvarchar](50) NULL,
	[delta] [nvarchar](50) NULL,
	[rSigma] [nvarchar](50) NULL,
	[valid] [nvarchar](50) NULL,
	[signal] [nvarchar](50) NULL,
	[lag] [nvarchar](50) NULL
) ON [PRIMARY]
GO



CREATE TABLE [dbo].[CalculationResult](
	[PRODUKTCODE] [nvarchar](50) NULL,
	[TREND_WERT] [nvarchar](50) NULL,
	[TREND_WERT_2] [nvarchar](50) NULL,
	[CL] [nvarchar](50) NULL,
	[LCL] [nvarchar](50) NULL,
	[UCL] [nvarchar](50) NULL,
	[LAL] [nvarchar](50) NULL,
	[UAL] [nvarchar](50) NULL,
	[TS_ABS] [nvarchar](50) NULL,
	[SORT_DATE] [nvarchar](50) NULL,
	[LAUFNR] [nvarchar](50) NULL,
	[EXCURSION] [nvarchar](50) NULL,
	[VIRT_OZID] [nvarchar](50) NULL,
	[VALUE] [nvarchar](50) NULL,
	[BatchID] [nvarchar](50) NULL,
	[lowSD] [nvarchar](50) NULL,
	[uppSD] [nvarchar](50) NULL,
	[mu] [nvarchar](50) NULL,
	[sigma] [nvarchar](50) NULL,
	[upp] [nvarchar](50) NULL,
	[delta] [nvarchar](50) NULL,
	[rSigma] [nvarchar](50) NULL,
	[valid] [nvarchar](50) NULL,
	[signal] [nvarchar](50) NULL,
	[lag] [nvarchar](50) NULL,
	[Column1] [nvarchar](50) NULL,
	[Column2] [nvarchar](50) NULL
) ON [PRIMARY]
GO




CREATE TABLE [dbo].[Calculations](
	[CalcID] [int] NOT NULL,
	[CalcDate] [datetime] NULL,
	[Num_VIRT_OZID_total] [int] NULL,
	[Num_VIRT_OZID_fit_stat] [int] NULL,
	[Percent_VIRT_OZID_fit_stat] [real] NULL,
	[Num_VIRT_OZID_not_fit_stat] [int] NULL,
	[UserId] [int] NULL,
	[Timepoint_of_RawData] [datetime] NULL,
	[Ver_R_mod] [nvarchar](50) NULL,
	[R_Errors] [nvarchar](max) NULL,
	[Note_from_user] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[DataGrid](
	[ID] [nvarchar](50) NULL,
	[ProductCode] [nvarchar](50) NULL
) ON [PRIMARY]
GO




CREATE TABLE [dbo].[Graphs](
	[GraphName] [nvarchar](max) NULL,
	[VIRT_OZID] [nvarchar](50) NULL,
	[ImageValue] [image] NULL,
	[CalcID] [nvarchar](50) NULL,
	[ID] [nvarchar](250) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[Persons](
	[ID] [int] NOT NULL,
	[Nachname] [nvarchar](max) NULL,
	[Vorname] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[Products](
	[ID] [int] NULL,
	[PRODUKTCODE] [nvarchar](250) NULL,
	[SORT_DATE] [datetime] NULL,
	[TS_ABS] [datetime] NULL,
	[LAUFNR] [nvarchar](50) NULL,
	[CHNR_ENDPRODUKT] [nvarchar](max) NULL,
	[PROCESS_CODE] [nvarchar](max) NULL,
	[PROCESS_CODE_NAME] [nvarchar](max) NULL,
	[PARAMETER_NAME] [nvarchar](max) NULL,
	[ASSAY] [nvarchar](max) NULL,
	[VIRT_OZID] [nvarchar](max) NULL,
	[TREND_WERT] [nvarchar](250) NULL,
	[TREND_WERT_2] [nvarchar](250) NULL,
	[ISTWERT_LIMS] [nvarchar](250) NULL,
	[LCL] [nvarchar](250) NULL,
	[UCL] [nvarchar](250) NULL,
	[CL] [nvarchar](250) NULL,
	[UAL] [nvarchar](250) NULL,
	[LAL] [nvarchar](250) NULL,
	[DECIMAL_PLACES_XCL_SUBSTITUTED] [nvarchar](250) NULL,
	[DECIMAL_PLACES_AL] [nvarchar](250) NULL,
	[DATA_TYPE] [nvarchar](max) NULL,
	[SOURCE_SYSTEM] [nvarchar](max) NULL,
	[EXCURSION] [nvarchar](10) NULL,
	[REFERENCED_CPV] [nvarchar](max) NULL,
	[IS_IN_RUN_NUMBER_RANGE] [nvarchar](max) NULL,
	[LOCATION] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[ProductsFiltered](
	[ID] [nvarchar](250) NULL,
	[PRODUKTCODE] [nvarchar](250) NULL,
	[SORT_DATE] [nvarchar](250) NULL,
	[TS_ABS] [nvarchar](250) NULL,
	[LAUFNR] [nvarchar](max) NULL,
	[CHNR_ENDPRODUKT] [nvarchar](max) NULL,
	[PROCESS_CODE] [nvarchar](max) NULL,
	[PROCESS_CODE_NAME] [nvarchar](max) NULL,
	[PARAMETER_NAME] [nvarchar](max) NULL,
	[ASSAY] [nvarchar](max) NULL,
	[VIRT_OZID] [nvarchar](max) NULL,
	[TREND_WERT] [nvarchar](250) NULL,
	[TREND_WERT_2] [nvarchar](250) NULL,
	[ISTWERT_LIMS] [nvarchar](250) NULL,
	[LCL] [nvarchar](250) NULL,
	[UCL] [nvarchar](250) NULL,
	[CL] [nvarchar](250) NULL,
	[UAL] [nvarchar](250) NULL,
	[LAL] [nvarchar](250) NULL,
	[DECIMAL_PLACES_XCL_SUBSTITUTED] [nvarchar](250) NULL,
	[DECIMAL_PLACES_AL] [nvarchar](250) NULL,
	[DATA_TYPE] [nvarchar](max) NULL,
	[SOURCE_SYSTEM] [nvarchar](max) NULL,
	[EXCURSION] [nvarchar](10) NULL,
	[REFERENCED_CPV] [nvarchar](max) NULL,
	[IS_IN_RUN_NUMBER_RANGE] [nvarchar](max) NULL,
	[LOCATION] [nvarchar](max) NULL,
	[UserID] [nvarchar](50) NULL,
	[ModifiedDate] [nvarchar](50) NULL,
	[GraphID] [nvarchar](50) NULL,
	[CalcID] [nvarchar](50) NULL,
	[FilterID] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[ProductsFilteredTemp](
	[ID] [nvarchar](250) NULL,
	[PRODUKTCODE] [nvarchar](250) NULL,
	[SORT_DATE] [nvarchar](250) NULL,
	[TS_ABS] [nvarchar](250) NULL,
	[LAUFNR] [nvarchar](max) NULL,
	[CHNR_ENDPRODUKT] [nvarchar](max) NULL,
	[PROCESS_CODE] [nvarchar](max) NULL,
	[PROCESS_CODE_NAME] [nvarchar](max) NULL,
	[PARAMETER_NAME] [nvarchar](max) NULL,
	[ASSAY] [nvarchar](max) NULL,
	[VIRT_OZID] [nvarchar](max) NULL,
	[TREND_WERT] [nvarchar](250) NULL,
	[TREND_WERT_2] [nvarchar](250) NULL,
	[ISTWERT_LIMS] [nvarchar](250) NULL,
	[LCL] [nvarchar](250) NULL,
	[UCL] [nvarchar](250) NULL,
	[CL] [nvarchar](250) NULL,
	[UAL] [nvarchar](250) NULL,
	[LAL] [nvarchar](250) NULL,
	[DECIMAL_PLACES_XCL_SUBSTITUTED] [nvarchar](250) NULL,
	[DECIMAL_PLACES_AL] [nvarchar](250) NULL,
	[DATA_TYPE] [nvarchar](max) NULL,
	[SOURCE_SYSTEM] [nvarchar](max) NULL,
	[EXCURSION] [nvarchar](10) NULL,
	[REFERENCED_CPV] [nvarchar](max) NULL,
	[IS_IN_RUN_NUMBER_RANGE] [nvarchar](max) NULL,
	[LOCATION] [nvarchar](max) NULL,
	[UserID] [nvarchar](50) NULL,
	[ModifiedDate] [nvarchar](50) NULL,
	[GraphID] [nvarchar](50) NULL,
	[CalcID] [nvarchar](50) NULL,
	[FilterID] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[ProdUpdated](
	[ID] [int] NULL,
	[PRODUKTCODE] [nvarchar](250) NULL,
	[SORT_DATE] [datetime] NULL,
	[TS_ABS] [datetime] NULL,
	[LAUFNR] [nvarchar](50) NULL,
	[CHNR_ENDPRODUKT] [nvarchar](max) NULL,
	[PROCESS_CODE] [nvarchar](max) NULL,
	[PROCESS_CODE_NAME] [nvarchar](max) NULL,
	[PARAMETER_NAME] [nvarchar](max) NULL,
	[ASSAY] [nvarchar](max) NULL,
	[VIRT_OZID] [nvarchar](max) NULL,
	[TREND_WERT] [nvarchar](250) NULL,
	[TREND_WERT_2] [nvarchar](250) NULL,
	[ISTWERT_LIMS] [nvarchar](250) NULL,
	[LCL] [nvarchar](250) NULL,
	[UCL] [nvarchar](250) NULL,
	[CL] [nvarchar](250) NULL,
	[UAL] [nvarchar](250) NULL,
	[LAL] [nvarchar](250) NULL,
	[DECIMAL_PLACES_XCL_SUBSTITUTED] [nvarchar](250) NULL,
	[DECIMAL_PLACES_AL] [nvarchar](250) NULL,
	[DATA_TYPE] [nvarchar](max) NULL,
	[SOURCE_SYSTEM] [nvarchar](max) NULL,
	[EXCURSION] [nvarchar](10) NULL,
	[REFERENCED_CPV] [nvarchar](max) NULL,
	[IS_IN_RUN_NUMBER_RANGE] [nvarchar](max) NULL,
	[LOCATION] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[TempParams](
	[inputfile] [nvarchar](max) NULL,
	[productcode] [nvarchar](max) NULL,
	[outputfolder] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO




CREATE TABLE [dbo].[VIRT_OZID_per_calculation](
	[VIRT_OZID] [nvarchar](50) NULL,
	[CalcID] [nvarchar](50) NULL,
	[TotalN] [nvarchar](50) NULL,
	[Additional_note] [nvarchar](max) NULL,
	[FitStatistically] [nvarchar](50) NULL,
	[RelevantForDiscussion] [nvarchar](50) NULL,
	[GraphID] [nvarchar](50) NULL,
	[KPI0] [nvarchar](50) NULL,
	[KPI1] [nvarchar](50) NULL,
	[KPI2] [nvarchar](50) NULL,
	[KPI3] [nvarchar](50) NULL,
	[Active] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO









/****** Object:  Index [Index_CalcID]    Script Date: 18.03.2023 23:20:36 ******/
CREATE NONCLUSTERED INDEX [Index_CalcID] ON [dbo].[CalculationRaw]
(
	[CalcID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO



/****** Object:  Index [IX_CalculationRaw_VIRT_OZID]    Script Date: 18.03.2023 23:20:47 ******/
CREATE CLUSTERED INDEX [IX_CalculationRaw_VIRT_OZID] ON [dbo].[CalculationRaw]
(
	[VIRT_OZID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO


GO

/****** Object:  Index [ClusteredIndex-20221124-174814]    Script Date: 18.03.2023 23:22:15 ******/
CREATE CLUSTERED INDEX [ClusteredIndex-20221124-174814] ON [dbo].[Products]
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO




CREATE PROCEDURE [dbo].[dataIn] 
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    insert into ProductsFiltered
	select * from ProductsFilteredTemp
	
END
GO






CREATE PROCEDURE [dbo].[InsertIntoCalculation] 
	@calcID as nvarchar(50), @date as nvarchar(50)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

   insert into CalculationRaw select @calcID,@date,[ProduktCode],[TREND_WERT],[TREND_WERT_2],[CL],[LCL],[UCL],[LAL],[UAL],[TS_ABS],[SORT_DATE],[LAUFNR],[EXCURSION],[VIRT_OZID],[VALUE],[BatchID],[lowSD],[uppSD],[mu],[sigma],[upp],[delta],[rSigma],[valid],[signal],[lag] from CalculationResult
END
GO




CREATE procedure [dbo].[nPercentStatisticallyFit]
	-- Add the parameters for the stored procedure here
	@CalcId nvarchar(100)
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	declare @a float
declare @b float
declare @c float
set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '0')
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId )
set @c = (@a/@b)*100
return (@a/@b)*100
END
GO





CREATE PROCEDURE [dbo].[ProdUpdate]  
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;


delete from ProdUpdated
insert into ProdUpdated
select distinct * from Products 

delete from Products

insert into Products
select distinct * from ProdUpdated  


END
GO




CREATE FUNCTION [dbo].[KPIcount0] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS int
AS
BEGIN
	-- Declare the return variable here
	declare @a int


set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '0' and VIRT_OZID = @VIRT_OZID)

return @a
END

GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[KPIcount1] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS int
AS
BEGIN
	-- Declare the return variable here
	declare @a int


set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '1' and VIRT_OZID = @VIRT_OZID)

return @a
END
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[KPIcount2] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS int
AS
BEGIN
	-- Declare the return variable here
	declare @a int


set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '2' and VIRT_OZID = @VIRT_OZID)

return @a
END
GO



-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date, ,>
-- Description:	<Description, ,>
-- =============================================
CREATE FUNCTION [dbo].[KPIcount3] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS int
AS
BEGIN
	-- Declare the return variable here
	declare @a int


set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '3' and VIRT_OZID = @VIRT_OZID)

return @a
END
GO








CREATE FUNCTION [dbo].[notStatisticallyFit] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100)
)
RETURNS int
AS
BEGIN
	-- Declare the return variable here
	declare @a float
    declare @b float
	declare @c float
    declare @d float
	declare @e float

--set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '0')
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '1')
set @c = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '2')
set @d = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '3')

return (@b+@c+@d)

END

GO






CREATE FUNCTION [dbo].[PercentStatisticallyFit] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100)
)
RETURNS float
AS
BEGIN
	-- Declare the return variable here
	declare @a float
declare @b float

set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '0')
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId )

return (@a/@b)*100

END
GO





CREATE FUNCTION [dbo].[PercentStatisticallyFit0_Ozid] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS float
AS
BEGIN
	-- Declare the return variable here
	declare @a float
declare @b float
declare @c float

set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '0' and VIRT_OZID = @VIRT_OZID)
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and VIRT_OZID = @VIRT_OZID)

if (@b=0)  select @c = 0
else  select @c = (@a/@b)*100
return cast(@c as decimal(10,2)) 
END
GO




create FUNCTION [dbo].[PercentStatisticallyFit1] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100)
)
RETURNS float
AS
BEGIN
	-- Declare the return variable here
	declare @a float
declare @b float

set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '1')
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId )

return (@a/@b)*100

END
--select [dbo].[PercentStatisticallyFit] ('05112022220840')
GO




CREATE FUNCTION [dbo].[PercentStatisticallyFit1_Ozid] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS float
AS
BEGIN
	-- Declare the return variable here
	declare @a float
declare @b float
declare @c float

set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '1' and VIRT_OZID = @VIRT_OZID)
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and VIRT_OZID = @VIRT_OZID)

if (@b=0)  select @c = 0
else  select @c = (@a/@b)*100
return cast(@c as decimal(10,2))  
END
GO




CREATE FUNCTION [dbo].[PercentStatisticallyFit2_Ozid] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS float
AS
BEGIN
	-- Declare the return variable here
	declare @a float
declare @b float
declare @c float

set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '2' and VIRT_OZID = @VIRT_OZID)
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and VIRT_OZID = @VIRT_OZID)

if (@b=0)  select @c = 0
else  select @c = (@a/@b)*100
return cast(@c as decimal(10,2))  
END
GO




CREATE FUNCTION [dbo].[PercentStatisticallyFit3_Ozid] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100), @VIRT_OZID nvarchar(100)
)
RETURNS float
AS
BEGIN
	-- Declare the return variable here
	declare @a float
declare @b float
declare @c float

set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '3' and VIRT_OZID = @VIRT_OZID)
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and VIRT_OZID = @VIRT_OZID)

if (@b=0)  select @c = 0
else  select @c = (@a/@b)*100
return cast(@c as decimal(10,2))  
END
GO




CREATE FUNCTION [dbo].[PerStatisticallyFit] 
(
	-- Add the parameters for the function here
	@CalcId nvarchar(100)
)
RETURNS float
AS
BEGIN
	-- Declare the return variable here
	declare @a float
    declare @b float
	declare @c float
    declare @d float
	declare @e float
set @a = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '0')
set @b = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '1')
set @c = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '2')
set @d = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId and signal = '3')

set @e = (select distinct count(signal)  from CalculationRaw where CalcID = @CalcId )

return  cast(((@a)/@e) *100 as decimal(10,0))  

END

GO




