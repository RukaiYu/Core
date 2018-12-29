## Quick Start

### Contents

- Host program
- Design layout
- Write code and run
- Function calls and event response

### Host program

To complete a demonstration of Tangram, we need to prepare a **Host Program** in advance. In a production environment, the **Host Program** may be a program under development or a third-party program. But to be as simple as possible, we create a new project now.  

Open Visual Studio and select File > New > Project. In Add Project, choose Visual C++ > MFC > MFC Application. We will name this project SimpleMFCApplication.

![Capture2](/Assets/Capture2.png)

In the Wizard, we chose to create an SDI window.

![Capture3](/Assets/Capture3.png)

Remove the status bar

![Capture5](/Assets/Capture5.png)

To simplify the code, remove all advanced features

![Capture6](/Assets/Capture6.png)

Click Finish to complete the wizard. Build and run this MFC program, we will see

![Capture8](/Assets/Capture8.png)

At this point, we inspect the program through the Spy++ tool and we will see two nested windows

![Capture9](/Assets/Capture9.png)

This window structure just meets the requirements of Tangram.

### Design layout

Here are two kinds of elements to be added to this program.

- .NET control - System.Windows.Forms.PropertyGrid
- Internet Explorer control

We plan to divide the above window into two parts. Place the .NET control on the left and the browser control on the right.  
To do this, we need to write a layout file.

    <tangram>
      <window>
        <node id='splitter' name='view' rows='1' cols='2' width='300,300' height='100' middlecolor='RGB(255,255,255)'>
          <node name='node1' id='CLRCtrl' cnnid='System.Windows.Forms.PropertyGrid,tangram' caption='node1'/>
          <node name='hostview' caption='host' id='hostview' />
        </node>
      </window>
    </tangram>

This is a layout file for Tangram. It is used to tell Tangram how to load and place elements correctly. You can read the [Layout Guide](/Docs/Layout.md) for detailed syntax. But now just copy and save it to the executable directory. We will name this file **SimpleMFCApplication.xml**.

### Write code and run

To work with Tangram, we add the Tangram.h header file first.

    // MainFrm.cpp : implementation of the CMainFrame class
    //

    #include "stdafx.h"
    #include "SimpleMFCApplication.h"

    #include "MainFrm.h"

    #include "Tangram.h"

We add the code to the CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct) method. This is not absolute. Just because you can easily get the window handle.

Initialize COM first.

	// COM initialization
    ::OleInitialize(nullptr);

Load the Tangram COM object.

	// Get the global Tangram object
	CComPtr<ITangram> ppTangram;
	ppTangram.CoCreateInstance(L"tangram.tangram");
	ITangram* pTangram = ppTangram.Detach();

Create WndPage and WndFrame objects. Load the SimpleMFCApplication.xml file for extension. To learn more about WndPage and WndFrame, read our [Programming Guide](/Docs/Programming_Guide.md).

	if (pTangram)
	{
		// Create WndPage
		CComPtr<IWndPage> pPage;
		pTangram->CreateWndPage((LONGLONG)m_hWnd, &pPage);
		if (pPage)
		{
			// Create WndFrame
			CComPtr<IWndFrame> pFrame;
			pPage->CreateFrame(CComVariant((LONGLONG)m_hWnd),
				CComVariant((LONGLONG)m_wndView.m_hWnd), BSTR(L"First Frame"), &pFrame);
			if (pFrame) {
				// Extend Window
				CComPtr<IWndNode> pNode;
				pFrame->Extend(BSTR(L"First Node"), BSTR(L"SimpleMFCApplication.xml"), &pNode);
			}
		}
	}

Build and run again, you will see the effect after the layout is in effect.

![Capture10](/Assets/Capture10.png)

Use Spy++ to inspect the window again.

[1]

You can try modifying the SimpleMFCApplication.xml. Open the program again and you will see the difference.
