# CsharpActiveX
Simple project for creating a C# ActiveX and Container. It Provides a complete example of integrating with ActiveX, IE and C#.

This is the source code and Visual Studio project behind the Stackoverflow question
[How to connect a C# ActiveX event handler in Javascript](http://stackoverflow.com/questions/30311383/how-to-connect-a-c-sharp-activex-event-handler-in-javascript)

It combines three code examples:

* [WebBrowser.ObjectForScripting Property](https://msdn.microsoft.com/en-us/library/system.windows.forms.webbrowser.objectforscripting.aspx)
* [How to develop and deploy ActiveX control in C#](http://blogs.msdn.com/b/asiatech/archive/2011/12/05/how-to-develop-and-deploy-activex-control-in-c.aspx)
* [How can I make an ActiveX control written with C# raise events in JavaScript when clicked?](http://stackoverflow.com/questions/1455577/how-can-i-make-an-activex-control-written-with-c-sharp-raise-events-in-javascrip)

## Things learned

* Don't use document.write inside the event handler as the object will be cleared from the page.
* ActiveX object utilized by Javascript needs to be obtained from javascript.
