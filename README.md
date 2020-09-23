<div align="center">

## Article \#1Introduction to the Win32API


</div>

### Description

The API programming series is a set of articles dealing with a common theme: API programming in Visual Basic. In this article we’ll look at the concept of API as applied to Win32 programming.

In this article we’ll look at the concept of API as applied to Win32 programming.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sreejath S\. Warrier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sreejath-s-warrier.md)
**Level**          |Beginner
**User Rating**    |4.4 (44 globes from 10 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sreejath-s-warrier-article-1introduction-to-the-win32api__1-33357/archive/master.zip)





### Source Code

<B>Foreword</B><P>
The API programming series is a set of articles dealing with a common theme: API programming in Visual Basic. Though there are no hard and fast rules regarding the content of these articles, generally one article can be expected to contain issues related to API programming, explanation of one or more API calls with generously commented code snippets or bug reports. Depending on the subject, these code samples may expand to become a full-fledged application. <P>
In this article we’ll look at the concept of API as applied to Win32 programming.<P>
<B>Note:</B> In this article we’ll look at the concept of API as applied to Win32 programming.
<P><B>So what’s API</B>
<P>You may be a VB programmer, a C++ programmer, a Delphi programmer or even a C programmer. But the fact is if you’ve ever developed an application for the Windows platform then you have used the Win32 API, at least indirectly. Because, quite simply, any program you write for windows uses the Windows API. Each and every line of code you write is translated into corresponding API calls which the system uses to get the tasks done.<P>
API (Application Programming Interface) is a set of predefined Windows functions used to control the appearance and behaviour of every Windows element (from the outlook of the desktop window to the allocation of memory for a new process). Between them, these functions encapsulate the entire functionality of the Windows environment. So we can consider API as the native code of Windows. The other languages act as an attractive and often user-friendlier shell to the API promoting easier and automated access to it. An example is VB, which has replaced a sizeable portion of the API with its own functions. But every line of code written in VB is converted to its equivalent API calls.
<P>
<B>Where does the API reside?</B>
<P>The bulk of the API functions are encapsulated in a set of DLLs - kernel32.dll, user32.dll, gdi32.dll, shell32.dll, etc. These DLLs form the core of the Windows OS and are present on all windows machines (At least the ones that are bootable J). So unlike when you access a function from a third-party DLL, you don’t have to ship anything with your product when you use an API function. The effect this has on the distribution size and speed of your application has to be seen to be believed.<P>
<B>How can we access API functions from VB?</B>
<P>Since these functions are encapsulated within external DLLs they are not available in VB by default. Before invoking (i.e. using) them, we must declare them. Usually this is done in the General | Declarations part of the module in which they are to be used. One important thing should be remembered while declaring API functions within a form module. VB does not allow public Declare statements within form modules. So the Declare statements should be explicitly specified as Private since all object module members are considered Public by default. But this makes them invisible outside the module in which they are defined. So the functions should be declared in every module that needs to use them. Declaring them in a separate .BAS module where they can be declared as Public and are therefore accessible to all the modules eliminates this duplication. This also permits easier maintenance.
<P><B>Why use the API when I can achieve most of the things I want in VB?</B>
<P>First of all the features offered by (pure) VB are pretty limited when compared to what can be achieved in windows. The recent additions have alleviated this problem somewhat, but the functionality offered by VB is still woefully inadequate in many areas, especially systems programming. Using the API can help you to work around this and would be much simpler than learning a new language.
<P>Secondly, due to its architecture, programs written in VB run slower than their C++ or Delphi counterparts. This is the price one has to pay for easier debugging, simpler coding and a fast development cycle. API being the native interface to Windows is much faster than VB when compared to VB code and since the support files are present in all Windows installations (more about this in the Where does the API reside section) the distribution size is smaller too. So using the API in VB allows one to have the best of both worlds.
<P>
<B>So why use VB at all?</B>
<P>I guess I went a bit overboard in praising the API. Because, for all its advantages, API programming does have some critical disadvantages. It is much more tougher and much less forgiving of programmer errors. Also it is child’s play to crash your application or your system with a malformed API call. And certain API calls when (im)properly used can even render your system unbootable or worse. <P>Scared? There is no need to be. Just try to understand what you’re doing and how to do it properly, exercise mature prudence and lo! You’ve got yourself a slimmer, faster application doing things that’d make ordinary VB programmers go green with envy. Good luck!<P>
Copyright © 2001 Sreejath S. Warrier
<P>

