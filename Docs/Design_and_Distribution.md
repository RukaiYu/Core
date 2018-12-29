## Design and Distribution

### Contents

- Structure
- Layout Persistence
- Looking for Extension Points
- Extension
- Visual Design
- Distribution and Redesign

### Structure

Tangram defines a tree-level structure on the interface layer. Developers can declare that some of the top-level windows in the program are allowed to contain extensible child windows. Tangram calls this type of top-level window called **Page**. These extensible child windows are called **Frame**. Each **Frame** window can be extended into a grid structure, and each cell in the grid is called a **Node**.

### Layout Persistence

Every program developed based on Tangram has a configuration file. Includes all the extended information for the program. This file is usually generated automatically by Tangram. Located in the `C:\Program Data\Tangram` directory. Developers can update configuration information use programmatically or visual design toolbox. During the release, the developer needs to contains the configuration file in the installation package and install it into the executable directory. The executable program is extended at the startup by parsing the configuration.

### Looking for Extension Points

Tangram uses the following API to define the specified window handle as a **Page** or **Frame**.

    HRESULT STDMETHODCALLTYPE CreateWndPage( 
            LONGLONG hWnd,
            /* [retval][out] */ IWndPage **ppWndPage)

Define a top-level window as a **Page**.

    HRESULT STDMETHODCALLTYPE CreateFrame( 
            VARIANT ParentObj,
            VARIANT HostWnd,
            BSTR bstrFrameName,
            /* [retval][out] */ IWndFrame **pRetFrame) = 0;

Define a child window as a **Frame**.

### Extension

Tangram uses

    HRESULT STDMETHODCALLTYPE Extend( 
            BSTR bstrKey,
            BSTR bstrXml,
            /* [retval][out] */ IWndNode **ppRetNode)

to extend the **Frame** window. Each key corresponds to a different layout. A **Frame** window can contain many layouts at the same time. And switch the layout according to the key. Developers also needs call

    pNode->put_SaveToConfigFile(true);

to make sure that all the layout information is saved to the configuration file.

### Visual Design

To simplify the work, Tangram allows use visual design toolbox to generate layout XML instead of writing them by hand. In order to wake up the design toolbox, you need to register a COM tab component, **TangramTabbedWnd.dll**.

    regsvr32 /s "{{Tangram Root Directory}}\Build\Lib\TangramTabbedWnd.dll"

Make sure your extension XML contains undesigned nodes.

    <node name='foo' caption='bar'/>

Run the program and you will see that the undesigned node becomes a gray window.

![Capture16](/Assets/Capture16.png)

Clicking on the gray window will bring up the visual design toolbox. Similar like this.

![Capture17](/Assets/Capture17.png)

With the visual design toolbox, developers can extend the program instantly. The extended results are immediately saved to the configuration file.

### Distribution and Redesign

During the release, the developer needs to contains the configuration file in the installation package and install it into the executable directory. When the program is first run on the client, Tangram will copy the configuration file to the corresponding `C:\Program Data\Tangram` directory. Then parse and start the extension. At this point, Tangram allows the user to reopen the visual design toolbox and modify the program layout. These changes are applied to the configuration file in the `C:\Program Data\Tangram` directory. Users can perform secondary development of the program in this way. We call it **Redesign**.


