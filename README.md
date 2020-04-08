PXL Handler Proof of Concept
============================

This demo registers a custom URI scheme `pxl:` and a program that will handle this scheme. The handler is called `PxlHandler.exe` and can be built with Visual Studio.

You must first run the program without any parameters to insert the correct entries into you registry. In subsequent invocations, you can run the program with the `pxl:`-based URI as a commandline parameter.

A small VueJS web application allows you to test the new URI type. Open a command window in the directory `pxl-link-demo` and run

```bat
  npm install
  npm run serve
```

Note that Chrome will ask you for a confirmation to follow the custom scheme. In an Electron app, this confirmation window can most likely be suppressed by configuring the built-in browser.

The current implementation of `PxlHandler.exe` shows a console window with messages. In a production situation you would suppress this window and let the application run invisibly.
