<div align="center">

## DTS Class Module and Events


</div>

### Description

The program Creates a VB Class Module from a DTS Package on a SQL Server, with all events, and its own events(Progress, currentTask, etc)

It creates a very compact script, so you can very large packages in a single routine.

The Class Module "ClassDTSScript" is what is created when you get a package from the server and script it. Simply remove the example, and add one you have done to test it.  The example execution asks you to navigate to the source and destination Access Databases, and uses the filepath to pass in an ADO style connection string for the source and destination connections. The parsing routine in the class module will work for SQL Server or ACCESS. I have not handled more complicated transformations such as Many to one column mappings and such, but Execute SQL and DATAPUMP Tasks work quite well. The Example "ClassDTSScript" module included was created from a package in SQL Server 7, and includes a couple of queries, two transformations, and running a stored procedure with a parameter, as well as demonstrating using the events that are called by the DTS Package object.  Read the comments in the code carefully to better understand the uses.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-12-31 10:58:34
**By**             |[Darryn Frost](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/darryn-frost.md)
**Level**          |Advanced
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[DTS\_Class\_4532812312001\.zip](https://github.com/Planet-Source-Code/darryn-frost-dts-class-module-and-events__1-30243/archive/master.zip)








