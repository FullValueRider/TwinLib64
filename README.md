# Twin64Lib
 
This is a project that started with me trying to put together some classes to help remove boilerplate from answers to advent of code problems.  

It started as a C# library, which was a bit scary, as at the time I'd never written any C# code.  

With the advent of twin basic I've reimplemented the C# code in VBA and added lots more functionality.  The introduction of twin Basic has allowed a simpler approach to providing an activex library for VBA use so I've moved the VBA code into twinBasic.

TwinLib64 currently uses vanilla VBA.  This allows round tripping between twinBasic and VBA environments.  This is useful as it allows the codebase to be inspected by Rubberduck.

This library is, of course, a never ending project
