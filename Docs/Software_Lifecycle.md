## Software Lifecycle

![Software_Lifecycle](/Assets/Software_Lifecycle.png)

Tangram redefines how desktop software is built. Before we do that, let's compare the tranditional way.

### Tranditional software build process

The first type of software is very simple. It does not require additional third-party technology. Developers create new projects through IDE or manually. Independently coding, debug, and release compiled binaries. Because it is simple. Such programs tend to package all the functionality in a single executable file.  

![Software_Lifecycle1](/Assets/Software_Lifecycle1.png)

A slightly more complexly software might need to reference some components developed by a third party. These components are usually a dynamic link library file or a jar package file. These components need to be distributed along with the executable file.  

![Software_Lifecycle2](/Assets/Software_Lifecycle2.png)

In a more complicated situation, developers may also use some off-the-shelf open source software or code base. This code is compiled into a part of the executable.

![Software_Lifecycle3](/Assets/Software_Lifecycle3.png)

### New software build process

But in the eyes of Tangram. The desktop software is composed of three parts. Executable files, shells and content components. An executable file is essential for any program. It serves as the entry point to the program and determines what is loaded later. Tangram expects the executable file to have minimal responsibility. But you must include the Tangram library. Then the developers needs to develop a shell. A shell can be thought of as a framework for a program. It usually contains common functions such as top-level windows, main menu, account mechanism, and message queue. Tangram loads these shells by configuration. You can even control different users to load different shells through authorization.  

Finally, Tangram wants developers to package each functional module in the software into spearate components. Depending on the language, it may be a COM component, a dynamic link library, a .NET library, a jar package, or a url. These components are installed into the shell by Tangram. These components can run in the same process. UI components developed in different languages can work together to build a single user interface.

![Software_Lifecycle4](/Assets/Software_Lifecycle4.png)

Both the shell and the components are loaded by configuration. Therefore, Tangram can easily add, remove and replace these shells and components at runtime. Since the shell becomes configurable, developers can even use a third-party software to act as a shell. e.g, the Eclipse Workbench.

![Software_Lifecycle5](/Assets/Software_Lifecycle5.png)

In addition, Tangram can also use some closed source software as a shell. Only developers need to develop a special adapter so that this kind of software can load the Tangram library. For example, Tangram can load the Tangram library into an Office program through an Office Add-in as an adapter.

![Software_Lifecycle6](/Assets/Software_Lifecycle6.png)

