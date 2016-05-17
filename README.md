# XSPMonitor
monitoring the flow of JSP, ASPX or ASP pages

XSPMonitor is an alternative to the MVC scheme for monitoring the control of flow for the application pages.
The control of page execution is made procedural using a proprietary language. The XSPMonitor components are:
- A script using tests, loops, recursive calls, calls for pages
- An interpreter for the script, called the "monitor"
- pages : ASP ASPX or JSP, depending on the target environment

Ahead of the basic instructions, the script may include other functions : optimisation and traces, instanciation of data objects, database transaction management, etc.

XSPMonitor script can be placed in a single file, providing a view of the whole application flow and used components and making the maintenance easier
