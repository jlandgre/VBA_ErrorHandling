This shares and serves as documentation for our VBA user messaging and error trapping code base for consulting work.

We developed (and hereby openly share) a VBA **ErrorHandling** Class and code structure for user messaging in Microsoft Excel/VBA applications. For us, this replaced a messy practice of sprinkling MsgBox statements and rambling, hard-to-find text strings throughout code. The Class' design lets our applications "fail gracefully" when problems occur. Our projects simply need to instance **ErrorHandling** as a global object in driver subsroutines. It then unobtrusively co-exists with whatever our simulation or data analysis application is doing and is there to share messages with our users when needed.

**Background**
We develop in Python too, but Microsoft Excel/VBA development allows consulting clients to personally interact with applications using the Microsoft Office installed base in their companies. For non-coding users and often even ourselves, we also value Excel as a user-facing front end --valuing its outstanding multi-tabbed and richly formatted user interface.

Given VBA's long history, error tracing code has certainly been developed by VBA experts over the years. However, given the often closed nature of consulting and in-house development work, we have never encountered open-source examples. Ours is easy to implement and meets our goal of allowing applications we develop to behave like a considerate human --providing users with useful guidance and advising them of problematic conditions and how to resolve them. By posting openly, we hope that additional improvement tips will surface and allow us to continue to improve.

Our goal to deliver great user interfaces to clients means that we need a VBA code base that goes beyond [this type of simplistic "crash the cymbals to stop the proceedings" error handling advice](https://stackoverflow.com/questions/1038006/what-are-some-good-patterns-for-vba-error-handling). While we have yet to extend our approach to Python, there seems to be a similar opportunity there --leading to an excellent and unobtrusive system that goes beyond how to use Try/Except statements for user messaging in an object-oriented Python architecture such as we prefer.

**Messaging and Error Handling Requirements**</br>
Here are the specific requirements and characteristics embodied by the **ErrorHandling** VBA class:
* The application needs to gracefully close things down and inform the user when a fatal error occurs. An application-specific message might be, "The tank capacity calculation gave a negative kg amount. That's non-valid but could occur if you mis-entered residence time. Valid residence times are 0 to 8.0 hours.")
* The application needs to advise the user with a dialog box when an interesting but non-fatal condition occurs. Such messages serve an important training purpose too. An example might be, "The online store simulation is showing greater than 2000 transactions per hour. That's theoretically possible, but may be unrealistic given your entered, site maintenance time of 1.2 hours per day. You can suppress this message in the future by selecting the 'Turn off transaction limit' check box in Settings."
* The application needs to take advantage of Excel's cell comment capability to highlight conditions needing attention in user inputs or simulation outputs. An example might be, "The value in this cell cannot be blank. Enter an integer between 0 and 63". Typically, such Excel cell comment messages are combined with also displaying a dialog box error message to direct the user to look at the comment.
* Error handling needs to work in a nested stack of single-action functions in our object-oriented code architecture. To report the error at the end of execution, we need for the error handling to relay routine-specific messages back to the user-initiated driver subroutines that the user originally triggers with buttons and menu commands. From there, the driver can decide how to report the condition.
* Error handling needs to provide technical error tracing diagnostics when unexpected VBA errors occur. This is critical to partnering with users to track down and stamp out bugs and unhandled exceptions that inevitably crop up when co-developing complex applications. The user may not understand such messages, but, for speedy debugging and resolution, they can convey them back to the developer along with file examples.
* Error handling should minimize code overhead and have simple implementation in an application's subs and functions.  We do not want error trapping to cause bloated code that interferes with testing and troubleshooing the application's use cases.
* For VBA applications that will be run on remote or virtual machines, we need capability to toggle an application's messaging to "Auto mode" to generate a log file of accumulated user messages instead of displaying dialog boxes that would halt the proceedings or go unseen in a remote environment

**VBA Error Handling Use Case Demos**</br>
The repository file **demo.xlsm** contains the **ErrorHandling** class and examples.
* The **Driver Example** button calls an example with a fatal error in the top-level subroutines --with error messaging in "user-facing" language
* Nested Example shows an example of an error message (and associated code) from a fatal error occurring two levels deep in called functions
* **Nested VBA Error** shows developer-facing messaging for error tracing of an "unexpected" fatal error. Its "Called by...Called by" tracing is a helpful start to tracking down such an error. Usually, we debug such errors by turning off error handling in the driver program (e.g. IsHandle=False when initializing **ErrorHandle**))
* **Warning Message Example** is a code example for a non-fatal warning message displayed by a dialog box
* **Cell Comment Example** is a code example of directing the user to an explanatory cell comment

The  **Errors_** worksheet contains all user messages for the application. Typically, this sheet will be **xlVeryHidden** from the user. Listed messages are indexed by function or subroutine name and by an integer Base index for each. Base (e.g. default local code = 0) also triggers handling "unexpected" VBA errors.

Because VBA lacks the ability to programmatically sense its location during execution, we define an **sLocn** string as a local Constant in each routine to allow error message lookup and stack tracing.  The code architecture minimize the code overhead needed in each routine --allowing tracking to be [relatively] unobtrusive in the application's code.  The examples show that it requires a couple of standard lines at the beginning and end of each Sub or Function.  To make code predictable, we generally err on the side of just including error handling in all but the simplest functions.

The **Errors_** sheet's convention of assigning a Base index for each routine means that, in the application's code, messages can be referenced by one-indexed integer like 1, 2 or 3. This avoids needing to worry about whether the error number is duplicated in another routine, The Base indices can be renumbered as new functions are added to a module.

**Conclusion and Additional Notes**</br>
Our system is an important feature of delivering thoughtful user interfaces for sometimes-complex applications and reflects a consulting philosophy that places high value on excellent user interface design. It reflects the hard-won human learning that, if error handling is simple to implement while coding, we won't delude ourselves with the mythological developer thinking of "I will implement error trapping and user messaging later."

J.D. Landgrebe
Data Delve Engineer LLC
