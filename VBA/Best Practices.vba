
'Option Explicit at the top

'Do not use Select/Activate too frequently. Try to avoid if possible
'You can refer to "VBA Reusable Snippets" for snippets doing the above

'Refresh all data connections at the start
'Refresh Pivot Cache instead of refreshing all pivots which are based on the same data source

'Screenupdating, Display alerts, Asktoupdatelinks, Enable events are disabled at the start
'One caveat, this also bypasses "Password prompt to access database", so keep them after the Password prompt

'Try to toggle calculation mode off/on while deleting rows, appending, refreshing data/pivots

'Use Clearcontents instead of deleting rows

'Keep formula in top one row and paste as values for cells below [Lesser the formulas, faster the wkbk]
'If the above data is used as data source for Pivot, you can also delete cells below keeping only 1 formula row _
'  as Pivot stores the data in Cache

'Make sure the Variables are defined keeping future developments in mind
'For example, define Lastrow as Long instead of Integer, so that there won't be a limit of 32,767 for it





