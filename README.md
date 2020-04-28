README COVID19excel

 ===================================
 
 DISCLAIMER 
 
  This has been slapped together pretty quickly and I make no guarantees of accuracy
  nor fitness for any commercial or medical purposes. I am not and have never been
  an epedemiologist - this is purely a recreational mathematical investigation of
  whether a simple mathematical logistic equation model is a better way to represent
  the COVID19 pandemic and see where it is going.
  
  In the early days of an epidemic, it is clear that the limit (asymptote)
  to the logistic growth is changing (increasing) day by day. Evidence has been collected
  in the workbook of this.
  
  ==========================

  An excel workbook with macros that get latest confirmed cases time series from John Hopkins
  Github
  
  loads into the worksheet name as per gcsSeriesConf
    
    JH data has one column per date, and an increasing number of rows
     per country or state with reported cases...
    
     need to process this into a time series running down the worksheet
     one row per date...
    
    there are two ways this is done.
	
	1) The Initial Approach
	
	 a) tblData
	 
	 the initial approach uses a table named tblData (an excel listobject) on
	 worksheet myData which supports automated solver calculation of logistic equation
	 fits to the data.

    myData worksheet has a large number of named ranges and a table named tblData. For
	each country implemented, there is a set of columns and named ranges created. This
	has now been automated in the AddNewCountry macro.
	
    tblData has 1 row per day of data
    
	there are numerous charts in separate sheets, one per country
    as its not easy to rearrange excel tables (listobjects) columns,
    without damaging formulas that depend on them,
    the columns might see haphazard in order
    
    tblData has some columns to the right on NSW test data that is
     not available in the JH datafeed. In first versions, these were
	 manually fed in on a daily basis. Later, a macro DailyNSWHealth
	 was created to scrape the NSW testing and local transmission cases
	 from the NSW Health website.
    
    The JH data is often not the latest available. The loading
    macro process will not overwrite any manually entered value that is
    larger than that read from JH.
    
    the myData worksheet has a large set of named ranges, and in particular
    dynamic named ranges that will expand as the rows of data accumulate.
    
	b) tblRuns
	
     in addition to tblData, there is a table tblRuns to the right which is
     also a listobject, consisting o a list of the parameters that the Solver process
     is estimating, and used in the logistic equation output (referred to in the
	 relevant country columns of tblData)
	 
	 tblRuns has one row per country (or state in case of NSW). There are also
	 two non-standard rows Tests, for estimating growth in testing in NSW 
	 and NSWUnknowns for tracking progression of the number of local known
	 and local unknown confirmed cases. Neither of these were published by
	 other states, and became available at later dates of the outbreak.
	 
	 c) Generic Country
	 
	 There are a set of defined names that allow any country to be selected
	 for retrieval into "Generic" columns of tblData, supported by special
	 generic row in tblRuns. There needs to be an entry made for the country
	 in tblPopulation (to give the population etc). 
	 
	 The names are
	 
	 GenCountry 	= set this to the required country (eg. Ireland)
	 GenState 		= set to a specific state if you wish to extract just one states date
	 GenStart		= the date preceding first day of outbreak (day last 0), or 22/02/2020
	                  whichever is later
					  
	 GenName 		=  concatentation of GenCountry and GenState (calculated)
	 GenMatch	    = the row of tblPopulation table where specified GenCountry found (calculated)
	 
	 c) tblHist
    
     Further to the right is a tblHist for pasting the daily solutions or
	 more frequent logistic solutions. These can be sorted into order
	 so that a graph can be made showing the progression of the logistic
	 limit as data is added to tblData -- only implemented for Italy
	 and NSW.
	 
	 In the range A1:G2 is a table of parameters that is used by the Solver
	 solution. The macro SolveAll sets the set of parameters to refer to the
	 same country-specific parameter in the tblRuns rows (et beta, delta, Po)
	 
	 The reason this is needed is that the solver constraints could not be 
	 set from VBA to refer to changing rows of tblRuns --- it was achieved
	 by setting named ranges in the A1:G2 table, and using VBA to change
	 which row of tblRuns they were set to refer to. It was found at least
	 when the number of days available was low, solver needed constraints in
	 some cases in order to solve correctly.
	 
	 d) tblPopulation
	 
	 There is a solveMsgs worksheet that contains a table of population data
	 and other data for determining whether the SolveAll process is run for the
	 country/state and or
	 
	Above the tblData and tblHist are various structures for supporting labelling
	of chart series, titles and axis titles
	
	e) GetJHConfirmedData
	
	 Gets latest confirmed cases time series from John Hopkins Github
    loads into the worksheet name as per gcsSeriesConf
    
    JH data has one column per date, and an increasing number of rows
     per country or state with reported cases...
    
     need to process this into a time series running down the worksheet
     one row per date...
    
     this is done in the table named tblData (an excel listobject)
    
    then processes into a myData worksheet int
    myData worksheet has a large number of named ranges
    and a table named tblData
    tblData has 1 row per day of data
    there are numerous charts in separate sheets, one per country
    as its not easy to rearrange excel tables (listobjects) columns,
    without damaging formulas that depend on them,
    the columns might see haphazard in order
    
    tblData has some columns to the right on NSW test data that is
     not available in the JH datafeed. These are manually fed in on
     a daily basis.
    
    as well, the JH data is often not the latest available. The loading
     macro process will not overwrite any manually entered value that is
     larger than that read from JH.
    
    the myData worksheet has a large set of named ranges, and in particular
    dynamic named ranges that will expand as the rows of data accumulate.
    
    
     in addition to tblData, consisting of two parts.
     the top part is a list of the parameters that the Solver process
     is estimating, and used in the logistic equation output
    
     below is is an area (growing) for pasting the daily solutions
     and then sorting them so that the growth of the asymptotes and
     other changes can be easily seen
	 
	 e) (i) ScrapceDailyNSWHealth
	 
	 GetJHConfirmedData macro also invokes a call to the ScrapeDailyNSWHealth,
	 which scans back from current day, looking for .aspx files on NSW Health
	 website in which the data updates are provided. The user is prompted
	 for the number of days to scan back.
	 
	 Web-page scraping was used because initially the datasets were not being
	 published otherwise.
	 
	 e) (ii) GetAusTweet
	 
	 This macro scrapes the pinned tweet on the https://twitter.com/COVID_Australia
	 twitter used page. This page is updated throughout the day with the
	 latest data published by Australian States. The WA data is not typically
	 published until after 9 pm AEST. SA, Tasmania and NT are also late in
	 the day with their stats updates.
	 
	 This is generally more reliable than waiting for the John Hopkins data.
	 
	
	 e) Sub AddNewCountry
	 
     adds a new country to the myData worksheet table tblData,
     including creating the defined names and formulas applicable to it.
     does not yet get the tblRuns row right
     you need to put the country details in tblPopulation first, above the
     row labelled with country Tests
     also insert a row in the named range Countries (but that is now superseded
     by tblPopulation.
	
	2) Second Approach
	
	A second method was developed to support adding new countries to graphs without the
	logistic equation complexity.
	
	This is based on taking the time series data and transposing it into a new worksheet
	TST (or DeathsTimeS) that are row date ordered
	
	The one big flaw in this was that at the 21st April, the John Hopkins data for Tasmania
	from 3rd April to 21st was flawed (incorredt, compared to reading from https://twitter.com/Covid19Tas
	that directly read Tas govt data). You have to correct this every time you retrieve a new days
	data. The first approach will not overwrite existing data in myData tblData.
	
	The second approach has also been used to create a worksheet DeathsTimeS with deaths data.
	
	There is a macro which can create per capita data columns from TST (or DeathsTimeS), and the
	defined name dynamic ranges needed to track/graph the cases from 100 per capita 
	confirmed cases or 100 deaths per capita.
	
	 TtranspSeries Macro
	 
     Copy and transposes the time-series data sheet which have dates in column order
     across row 1 into worksheet with row-order for dates. Handles cases and
     deaths - dont feel recoveries mean enough to be useful to plot.
    
    
     Select countries to create plottable columns and range names for graphs of
     cases since Solvmsgs!mincases or deaths since Solvmsgs!MinDeaths per capita
    
     A set of defined names are created with dynamic ranges for each country for supporting
     graphing of the data.
    
     the deaths csv has a column for population of each county in the USA as dropped
     from JH github
    
     each time it is run, new country (or state) columns can be added, all can be rebuilt,
     or you can just have the latest data pasted into the TST or DeathsTimeS worksheets.
    
	AddSeriesToCht1 Macro
    
     This allows user to select a chart (located on a sheet, not tested for embedded chart)
     then copy the SERIES function for a particular country or state to 
	 series for a new country or state, substituting the new country name in place
	 of the country/state named in the template SERIES function.
    
     the defined name dynamic ranges need to be set up first for this to work
    
	=SERIES(TST!VictoriaSeriesLabel,TST!Victoria100Days,TST!Victoria100Cases,3)
	
	 The required dynamic names are established by the TtranspSeries macro (for 
	 TSTS and DeathsTimeS) or the AddCountry macro for tblData in myData worksheet.
	 
	 