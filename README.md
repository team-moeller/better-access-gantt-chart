# Better-Access-Gantt-Chart
Gantt charts in Access with JavaScript

## Why Better Access Gantt Chart?

Microsoft Access offers various chart types. However, when it comes to project planning, one type of chart is missing: the Gantt chart. This project aims to fill that gap.

Numerous JavaScript libraries exist online that address this topic, some of which are freely available. In this project, we use [Frappe Gantt](https://github.com/frappe/gantt), an open-source JavaScript library.

Our data source is a table in Access containing the project tasks. From this, we create a corresponding Gantt chart using Frappe Gantt Chart and display it in the Edge browser control.

The solution is designed to be interactive. Changes made by the user in the Gantt chart are automatically reflected in the task table.

## You want to give it a try?
1. Download the [latest release](https://github.com/team-moeller/better-access-gantt-chart/releases/latest)
2. Unpack the files to a trusted folder
3. Run the database
4. Push the button: "Create Gantt Chart"

## How to integrate into your own database?
**1. Import of the class modules**

First, the class module "cls_Better_Access_Gantt_Chart" and all class modules with the name "cls_Gantt_*" must be imported from the demo database into your Access database.

**2. Insert edge browser control on form**

The second step is to add an edge browser control to display the chart on a form. It is best to give the control a meaningful name. This is required later in the VBA code. I like to use the name "ctlEdgeBrowser" for this.

**3. Insert code for generating the Gantt chart**

The best thing to do is to add another button. In the click event, paste the following code:

```vba
Dim myGantt As cls_Better_Access_Gantt_Chart  
Set myGantt = New cls_Better_Access_Gantt_Chart
Set myGantt.Control = Me.ctlEdgeBrowser
' Add chart properties.
' The code in frm_Demo provides an example and demonstrates the possibilities.  
myGantt.ShowGanttChart  
```
