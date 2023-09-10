# __dsLiveEd__

A **PowerQuery-like MSAccess Application** for managing automation tasks of data manipulation or direct-data manipulation with it's **Live Editor** interface. 


### __QUICK FEATURE OVERVIEW__

- __Modes of Use__
  - __Live Editor UI__ - *It can be embedded in your application as a project reference or opened as an external (standalone) application.*
  - __Headless mode__ - *Those list of tasks to be applied can be exported as a presset in JSON, which can be imported in your application and just generate the resulting output without having to involve any visible UI.*
  - __Automation__ - *Formerly known as `OLE Automation`, allows to programatically use `dsLiveEd` from another application without even having to include `ds3x` as a project reference.*
  - __Command line__ - Allows executing automation tasks from command line switches, no programming skills neded.
- __Immutability support__ - *Allowing to go back and forward within the resulting state of each and every single transformation task applied to data tables (Excel doesn't support immutability so it won't go back and fordward by just switching between Excel tasks but they can be edited anyway).*

