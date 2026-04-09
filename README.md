# Better-Access-Gantt-Chart
Gantt charts in Access with JavaScript

## Why Better Access Gantt Chart?

Microsoft Access offers various chart types. However, when it comes to project planning, one type of chart is missing: the Gantt chart. This project aims to fill that gap.

Numerous JavaScript libraries exist online that address this topic, some of which are freely available. In this project, we use [Frappe Gantt](https://github.com/frappe/gantt), an open-source JavaScript library.

Our data source is a table in Access containing the project tasks. From this, we create a corresponding Gantt chart using Frappe Gantt Chart and display it in the Edge browser control.

The solution is designed to be interactive. Changes made by the user in the Gantt chart are automatically reflected in the task table.
