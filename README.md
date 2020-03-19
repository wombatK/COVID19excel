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
    
     this is done in the table named tblData (an excel listobject) on
	 worksheet myData

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
    
    
    in addition to tblData, there is a table to the right which is
     not a listobject, consisting of two parts.
     the top part is a list of the parameters that the Solver process
     is estimating, and used in the logistic equation output
    
     below is is an area (growing) for pasting the daily solutions
     and then sorting them so that the growth of the asymptotes and
     other changes can be easily seen