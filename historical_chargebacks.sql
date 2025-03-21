-- 3 year historical data FY 17-18 thru FY 19-20
-- (New database begun 2020-09-01)
-- Intended to be merged with data from new database
-- Includes hours worked by Contractor Cost Centers
-- Includes hours worked at a zero chargeback rate


-- Constants Set-up
DECLARE @lastMonthEnd AS DATE;
DECLARE @rptStartDate AS DATE;
DECLARE @lastPeriod AS NVARCHAR(255);
DECLARE @fiscalYear AS CHAR(4);
DECLARE @OVHD AS CHAR(5);
DECLARE @OVERHEAD AS CHAR(8);
DECLARE @DEFAULT AS CHAR(7);
DECLARE @NOCC AS CHAR(3);
DECLARE @Y AS CHAR(3);
DECLARE @N AS CHAR(2);
DECLARE @PROJECT AS CHAR(7);
DECLARE @PRODUCT AS CHAR(7);
DECLARE @SINGLEPRODUCT AS CHAR(14);
DECLARE @CCENTER AS CHAR(11);
DECLARE @SUPPORT AS CHAR(8);
DECLARE @ITBOCC AS CHAR(3);  -- Contractor Cost Center

-- Object Codes
DECLARE @CapObjCode AS CHAR(4);
DECLARE @CapSaaSObjCode AS CHAR(4);
DECLARE @SuptObjCode AS CHAR(4);
DECLARE @ProjObjCode AS CHAR(4);
DECLARE @ContractorCapObjCode AS CHAR(4);
DECLARE @ContractorCapSaaSObjCode AS CHAR(4);
DECLARE @ContractorSuptObjCode AS CHAR(4);
DECLARE @ContractorProjObjCode AS CHAR(4);
DECLARE @SuptObjCodeDesc AS VARCHAR(16);
DECLARE @ProjObjCodeDesc AS VARCHAR(16);
DECLARE @CapObjCodeDesc AS VARCHAR(16);
DECLARE @CapSaaSObjCodeDesc AS VARCHAR(16);
DECLARE @ContractorSuptObjCodeDesc AS VARCHAR(17);
DECLARE @ContractorProjObjCodeDesc AS VARCHAR(17);
DECLARE @ContractorCapObjCodeDesc AS VARCHAR(17);
DECLARE @ContractorCapSaaSObjCodeDesc AS VARCHAR(16);


-- Set the end date to the last date of data
SET @lastMonthEnd = '2020-08-31';

-- Look up Fiscal Year of just-completed month
SET @fiscalYear =
    (SELECT FiscalYear
       FROM dbo.MSP_TimeByDay
      WHERE CAST(TimeByDay AS DATE) = @lastMonthEnd);

-- Look up Fiscal Period of just-completed month
SET @lastPeriod =
    (SELECT FiscalPeriodName
       FROM dbo.MSP_TimeByDay
      WHERE CAST(TimeByDay AS DATE) = @lastMonthEnd);

-- Setting report start date by grabbing start of month 3 years ago
SET @rptStartDate = '2017-09-01'


SET @OVHD = 'Ovhd';
SET @OVERHEAD = 'Overhead';
SET @DEFAULT = 'Default';
SET @NOCC = '000';
SET @Y = 'Yes';
SET @N = 'No';
SET @PROJECT = 'Project';
SET @PRODUCT = 'Product';
SET @SINGLEPRODUCT = 'Single Product';
SET @CCENTER = 'Cost Center';
SET @SUPPORT = 'Support';
SET @ITBOCC = '316'  -- Contractor Cost Center

-- Object Codes
SET @CapObjCode = '2791';
SET @CapSaaSObjCode = '2691';
SET @SuptObjCode = '8166';
SET @ProjObjCode = '8164';
SET @ProjObjCodeDesc = '8164 Project';
SET @SuptObjCodeDesc = '8166 Support';
SET @CapObjCodeDesc = '2791 Cap';
SET @CapSaaSObjCodeDesc = '2691 Cap SaaS'
SET @ContractorCapObjCode = '2790';
SET @ContractorCapSaaSObjCode = '2690';
SET @ContractorSuptObjCode = '8165';
SET @ContractorProjObjCode = '8165';
SET @ContractorCapObjCodeDesc = '2790 Cap';
SET @ContractorCapSaaSObjCodeDesc = '2690 Cap SaaS';
SET @ContractorSuptObjCodeDesc = '8165 Proj/Support';
SET @ContractorProjObjCodeDesc = '8165 Proj/Support';


---------------------------------------------
-- Start of main query  
SELECT CASE P.TimeClass
          WHEN @OVHD THEN @OVERHEAD
          ELSE P.TimeClass
       END AS [Work Type]
	 , P.TimeClass
	 , P.ProjectName AS [Raw Proj Name] -- For auditing/debugging
     , CASE P.TimeClass
          WHEN @CCENTER THEN 'Cost Center Support'
          WHEN @PRODUCT THEN 'Product Support'
          WHEN @SUPPORT THEN 'Product Support' 
          WHEN @SINGLEPRODUCT THEN 'Product Support'
          WHEN @OVHD THEN @OVERHEAD
          ELSE P.TimeClass
       END AS [Time Type]
     , TBD.FiscalYear AS [Fiscal Year]
     , @lastperiod AS [Thru Fiscal Period]
     , LEFT(RTRIM(DATENAME(MONTH, TBD.TimeByDay)),3)
       + ', ' + LTRIM(STR(YEAR(TBD.TimeByDay))) AS [Month]
     , 'Q' + CAST(TBD.FiscalQuarter AS CHAR(1)) AS [Fiscal Quarter]
     , TBD.FiscalPeriodName AS [Fiscal Period]
     , CASE P.TimeClass
          WHEN @PROJECT THEN REPLACE(P.ProjectName, '_', ' ')
          WHEN @SINGLEPRODUCT THEN REPLACE(P.ProjectName, '_', ' ')
          ELSE LTRIM(T.TaskName)
       END AS [Project]
	 , CASE P.[Cap Project]
	      WHEN 0 THEN @N
		  ELSE @Y
	   END AS [Cap Project]
     , CASE 
		  WHEN P.[Cap Asset Number(s)] LIKE '%SaaS%'
		  THEN @Y
		  ELSE @N
	   END AS [SaaS Cap]
     , CASE P.[Cap Project]
          WHEN 1 THEN
             CASE 
                WHEN ISNULL(P.[Cap Asset Number(s)],'') = '' THEN '(Pending)'
                ELSE ISNULL(P.[Cap Asset Number(s)],'')
             END
          ELSE ''
       END AS [Cap Assets]
     , P.[ProjectOwnerName] AS [Project Manager]
     , R.ResourceName AS [Staff Member]
     , T.TaskName AS Task
     , CASE P.TimeClass
          WHEN @PROJECT THEN ISNULL(P.[IT Product], '')
          WHEN @SINGLEPRODUCT THEN ISNULL(P.[IT Product], '')
          WHEN @SUPPORT THEN ISNULL(P.[IT Product], '')
          WHEN @PRODUCT THEN ISNULL(T.[Task Product],'')
          ELSE ''
       END AS [Product] 
     , CAST(TBD.TimeByDay AS DATE) AS [Date]
     , ISNULL(TA.Comment, '') AS [Comment]
     , RCR.Rate AS [Rate]
     , CAST(TA.[Hours] AS MONEY) AS [Hours]
     , CAST(RCR.Rate * TA.[HOURS] AS MONEY) AS [Cost]
     , CASE P.[Cap Project]
          WHEN 0 THEN @N
          ELSE
             CASE ISNULL(T.[Cap Task], @DEFAULT)
                WHEN @DEFAULT THEN @Y
                ELSE T.[Cap Task]
             END
       END AS [Cap Task]
        --Exp Hours
     , CAST(CASE      
               -- START P.[Cap Task]
                   CASE P.[Cap Project]
                      WHEN 0 THEN @N
                      ELSE
                         CASE ISNULL(T.[Cap Task], @DEFAULT)
                            WHEN @DEFAULT THEN @Y
                            ELSE T.[Cap Task]
                         END
                   END     
               -- END CASE P.[Cap Task]
               WHEN @N THEN TA.[Hours]
               ELSE 0
            END AS MONEY) AS [Exp Hours]
       --Exp Cost
     , CAST(CASE      
               -- START P.[Cap Task]
                   CASE P.[Cap Project]
                      WHEN 0 THEN @N
                      ELSE
                         CASE ISNULL(T.[Cap Task], @DEFAULT)
                            WHEN @DEFAULT THEN @Y
                            ELSE T.[Cap Task]
                         END
                   END     
               -- END CASE P.[Cap Task]
               WHEN @N THEN CAST (RCR.Rate * TA.[Hours] AS MONEY)
               ELSE 0
            END AS MONEY) AS [Exp Cost]
       --Cap Hours
     , CAST(CASE      
               -- START P.[Cap Task]
                   CASE P.[Cap Project]
                      WHEN 0 THEN @N
                      ELSE
                         CASE ISNULL(T.[Cap Task], @DEFAULT)
                            WHEN @DEFAULT THEN @Y
                            ELSE T.[Cap Task]
                         END
                   END     
               -- END CASE P.[Cap Task]
               WHEN @Y THEN TA.[Hours]
               ELSE 0
            END AS MONEY) [Cap Hours]
       --Cap Cost
     , CASE      
          -- START P.[Cap Task]
              CASE P.[Cap Project]
                 WHEN 0 THEN @N
                 ELSE
                    CASE ISNULL(T.[Cap Task], @DEFAULT)
                       WHEN @DEFAULT THEN @Y
                       ELSE T.[Cap Task]
                    END
              END     
          -- END CASE P.[Cap Task]
          WHEN @Y THEN CAST(RCR.Rate * TA.[Hours] AS MONEY)
          ELSE 0
       END AS [Cap Cost]
 
------------------------------------------------------
-- Set the Object Codes and Object Code Descriptions,
     , CASE      
          -- START P.[Cap Task]
              CASE P.[Cap Project]
                 WHEN 0 THEN @N
                 ELSE
                    CASE ISNULL(T.[Cap Task], @DEFAULT)
                       WHEN @DEFAULT THEN @Y
                       ELSE T.[Cap Task]
                    END
              END     
          -- END CASE P.[Cap Task]
         WHEN @Y THEN 
				  CASE
				     WHEN P.[Cap Asset Number(s)] LIKE '%SaaS%' 
					 THEN 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @CapSaaSObjCode
						   ELSE  @ContractorCapSaaSObjCode
					    END
					 ELSE 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @CapObjCode
						   ELSE  @ContractorCapObjCode
					    END
			      END
         ELSE
            CASE 
               -- START P.[Work Type]
                  CASE P.TimeClass
                     WHEN @OVHD THEN @OVERHEAD
                     ELSE P.TimeClass
                  END
                -- END CASE P.[Work Type]
                WHEN @PROJECT THEN 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @ProjObjCode
						   ELSE  @ContractorProjObjCode
					    END
                ELSE 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @SuptObjCode
						   ELSE  @ContractorSuptObjCode
					    END
            END
       END AS [Object Code]
       
     , CASE                
	      -- START P.[Cap Task]
              CASE P.[Cap Project]
                 WHEN 0 THEN @N
                 ELSE
                    CASE ISNULL(T.[Cap Task], @DEFAULT)
                       WHEN @DEFAULT THEN @Y
                       ELSE T.[Cap Task]
                    END
              END     
          -- END CASE P.[Cap Task]
         WHEN @Y THEN 
				  CASE
				     WHEN P.[Cap Asset Number(s)] LIKE '%SaaS%' 
					 THEN 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @CapSaaSObjCodeDesc
						   ELSE  @ContractorCapSaaSObjCodeDesc
					    END
					 ELSE 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @CapObjCodeDesc
						   ELSE  @ContractorCapObjCodeDesc
					    END
			      END
         ELSE
            CASE 
               -- START P.[Work Type]
                  CASE P.TimeClass
                     WHEN @OVHD THEN @OVERHEAD
                     ELSE P.TimeClass
                  END
                -- END CASE P.[Work Type]
                WHEN @PROJECT THEN 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @ProjObjCodeDesc
						   ELSE  @ContractorProjObjCodeDesc
					    END
                ELSE 					    
						CASE 
						   WHEN RCR.CC_CODE NOT LIKE 'C%' AND  RCR.CC_CODE <> @ITBOCC
						   THEN  @SuptObjCodeDesc
						   ELSE  @ContractorSuptObjCodeDesc
					    END
            END

       END AS [Object Code Desc]

------------------------------------------------------

     , CASE P.[Cap Complete]
          WHEN '1' THEN @Y
          ELSE @N
       END AS [CapComplete]
     , CAST(P.ProjectStartDate AS DATE) AS [ProjectStart]
     , CAST(P.ProjectFinishDate AS DATE) AS [ProjectFinish]
 
-- Pull Resource's OVERHEAD CC from date-driven
-- Resource CC Rate custom look-up table,
-- not from Assignment
     , CASE 
           -- START P.[Work Type]
              CASE P.TimeClass
                 WHEN @OVHD THEN @OVERHEAD
                 ELSE P.TimeClass
               END
           -- END CASE P.[Work Type]
          WHEN @OVERHEAD THEN RCR.[CC_Code]
          ELSE CC.[CC Code]
       END AS [Charged CC Code]
     , CASE  
           -- START P.[Work Type]
              CASE P.TimeClass
                 WHEN @OVHD THEN @OVERHEAD
                 ELSE P.TimeClass
               END
           -- END CASE P.[Work Type]
          WHEN @OVERHEAD THEN RCR.[CC_Code] + ' - ' + RCR.CC_Desc
          ELSE CC.[Cost Center]
       END AS [Charged CC]

     , RCR.[CC_Code] AS [Work CC CODE]
     , RCR.[CC_Code] + ' - ' + RCR.CC_Desc AS [Work CC]

  FROM


-- Timesheet Lines  TA 
-- with Actuals By Day
( 
	-- Timesheet Line
    SELECT TSA.TimeByDay
         , TSL.ProjectName AS [Project]
         , TSL.TaskName AS [Task]
         , TSL.ResourceName AS [Staff Member]
         , ROUND(TSA.ActualWorkBillable * 4, 0) / 4.0 AS [Hours] -- Rounds to nearest 1/4 hr (fixes error in rounding in PWA)
         , ISNULL(TSA.Comment, '') AS [Comment]
         , TSA.TimesheetLineUID
         , TSL.ProjectUID
         , TSL.TaskUID
         , TSL.ResourceUID

      FROM MSP_TimesheetLine_UserView TSL  -- The lines on a Timesheet

    -- Filter on necessary Projects to keep from reading all time entries from all Projects.
	-- We pull these Projects in again below, to make all Project columns available
	-- and avoid having to name them above.
   	INNER JOIN [dbo].[MSP_EpmProject_UserView] FP
	   ON TSL.ProjectUID = FP.ProjectUID
	  AND FP.ProjectFinishDate >= @rptStartDate	-- Projects with work within date parameters

    INNER JOIN MSP_TimesheetActual_OlapView TSA -- The daily entries on a Timesheet Line
       ON TSL.TimesheetLineUID = TSA.TimesheetLineUID
      AND TSA.ActualWorkBillable > 0
      AND TSA.TimeByDay >= @rptStartDate
      AND TSA.TimeByDay <= @lastMonthEnd

 ) AS TA

-- Time utility table  TBD
 INNER JOIN dbo.MSP_TimeByDay TBD
    ON TA.TimeByDay = TBD.TimeByDay

-- Task User View  T
 INNER JOIN dbo.MSP_EpmTask_UserView T
    ON TA.TaskUID = T.TaskUID

-- Resource User View  R 
-- (Task Resources)    
 INNER JOIN [dbo].[MSP_EpmResource_UserView] R
    ON TA.ResourceUID = R.ResourceUID
   AND R.ResourceType = 2
   AND R.ResourceIsGeneric = 0

-- Project User View  P 
 INNER JOIN [dbo].[MSP_EpmProject_UserView] P 
    ON TA.ProjectUID = P.ProjectUID

-- Resource User View  PM
-- (Project Manager)

 INNER JOIN [dbo].[MSP_EpmResource_UserView] PM
    ON P.ProjectOwnerResourceUID = PM.ResourceUID

-- TASB's Cost Center View  CC
 LEFT OUTER JOIN dbo.vwCostCenter CC
    ON CC.[CC Code] =
       CASE
            WHEN @OVHD = ISNULL(P.[TimeClass], '') THEN R.ResourceCostCenter
            WHEN @NOCC = ISNULL(T.[Task Cost Center], @NOCC) THEN P.[Project Cost Center]
            ELSE T.[Task Cost Center]
       END

 ---------------------------------------------------------------

-- RCR
-- Based on ResourceCCRates
-- Custom Lookup Table

-- Look up each Staff Member's Cost Center in effect on the day the work was performed.
-- NOTE: This Inner Join also filters out any Resources who don't charge back.

   INNER JOIN
    ( SELECT D.MemberValue AS [Date]
         , C.MemberValue AS [CC_CODE]
         , C.MemberDescription AS [CC_Desc]
         , R.MemberValue AS [Name]
         , CASE ISNUMERIC(R.MemberDescription)
            WHEN 1 THEN CAST(R.MemberDescription AS money)
            ELSE 999
           END AS [Rate]
       FROM dbo.MSPLT_ResourceCCRates_UserView D
      INNER JOIN dbo.MSPLT_ResourceCCRates_UserView C
         ON D.LookupMemberUID = C.ParentLookupMemberUID
      INNER JOIN dbo.MSPLT_ResourceCCRates_UserView R
         ON C.LookupMemberUID = R.ParentLookupMemberUID
      WHERE D.ParentLookupMemberUID IS NULL
        AND D.MemberFullValue IS NOT NULL
    ) AS RCR
   ON TA.[Staff Member] = RCR.[Name]
  AND RCR.Date =
          ( SELECT MAX(RCR2.[Date])
              FROM
             ( SELECT D.MemberValue AS [Date]
                  , C.MemberValue AS [CC_CODE]
                  , R.MemberValue AS [Name]
                  , R.MemberDescription AS [Rate]
                FROM dbo.MSPLT_ResourceCCRates_UserView D
            INNER JOIN dbo.MSPLT_ResourceCCRates_UserView C
                  ON D.LookupMemberUID = C.ParentLookupMemberUID
            INNER JOIN dbo.MSPLT_ResourceCCRates_UserView R
                  ON C.LookupMemberUID = R.ParentLookupMemberUID
                WHERE D.ParentLookupMemberUID IS NULL
                  AND D.MemberFullValue IS NOT NULL
             ) AS RCR2
             WHERE TA.[Staff Member] = RCR2.Name
               AND TBD.TimeByDay >= RCR2.[Date]
          )
