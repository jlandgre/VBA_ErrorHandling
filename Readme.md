This shares the VBA user messaging and error trapping code base for consulting work. Given VBA's long history, such code has certainly been developed by others over the years. However, given the often proprietary nature of consulting and in-house development work, we have never encountered open-source examples for VBA (or Python). Ours is useful, easy to implement and meets our goal of allowing applications we develop to behave like a considerate human --advising users of problematic conditions and how to resolve them.

Our goal is a user messaging and error trapping code system that goes beyond [this type of simplistic "crash the cymbals to stop the proceedings" instructions for handling errors](https://stackoverflow.com/questions/1038006/what-are-some-good-patterns-for-vba-error-handling). While we have yet to extend our approach to Python, there seems to be a similar opportunity there. Our VBA approach should translate well to go beyond simplistic Python Try/Except tutorials.

When we develop in VBA versus Python, it is because Excel and VBA allow consulting clients to personally interact with the simulation or data handling application using the Microsoft Office installed base in their company. We also use Excel to take advantage of its outstanding multi-tabbed and richly formatted user interface. In these situations, we

To address this need, we developed a VBA **ErrorHandling** class for user messaging. This replaced an all-too-common practice of sprinkling MsgBox statements and rambling text strings throughout code. That is familiar to anyone who has worked with VBA, but we found it to become an intractable mess. **ErrorHandling** can be instanced as a global object and unobtrusively co-exists with whatever the simulation or data analysis application is doing. Here are the specific requirements and characteristics of our system:
* Ability to gracefully close things down and inform the user when a fatal error occurs while an application is running. An example might be, "The tank capacity calculation came up with a negative kg amount. That could occur if you mis-entered the residence time")
* Ability to warn or advise the user with a dialog box when an interesting but non-fatal condition occurs. An example might be, "The plant operation simulation is showing line utilization of greater than 23 hours per day. That's possible, but may be unrealistic given the need for "
* Ability to take advantage of Excel's unique cell comment capability to flag conditions that need attention in user inputs or simulation outputs
* Ability to work in a stack of single-action functions in our object-oriented code architecture. We need for the error handling to cascade messages back to the driver subroutines initiated by user actions such as pushing buttons and making menu selections.
* Provide technical error tracing diagnostics when unexpected VBA errors occur. This is critical to partnering with users to track down and stamp out bugs and unhandled exceptions that inevitably crop up in complex applications
* Minimize error trapping code overhead and have simple implementation in an application's subs and functions
* For VBA applications that will be run on remote or virtual machines, be able to toggle an application's messaging to "Auto mode" to generate a log file of accumulated user messages instead of displaying dialog boxes that will halt the proceedings in a remote environment

The file demo.xlsm contains the ErrorHandling class and examples.
* The Driver Example button calls an example with a "fatal" error in the top-level subroutines
* The Nested Example button shows an example of a user-facing error message from a fatal error two levels deep in called functions
* The Nested VBA Example shows an example of technical error tracing of an "unexpected" fatal error. Its "Called by...Called by" tracing is a helpful start to tracking down such an error. Usually, we conduct the full diagnostics by turning off error handling in the driver program xxx
* Warning Message Example is a code example for a non-fatal warning message
* Cell Comment Example is a code example of directing the user to an explanatory cell comment

The **Errors_** worksheet contains all user messages for the "application." These are indexed by a Base rows for each subroutine or function. In the application's functions, messages can be referenced by an integer like 1, 2 or 3. The ErrorHandling code ties this to the Base index such as 100 to generate the lookup key for the Driver Example. This makes it easier to assign new messages without having to constantly reference the **Errors_** table.

J.D. Landgrebe
Data Delve Engineer LLC
