package springswtword;

import com.sun.glass.ui.MenuBar;
import org.eclipse.jface.window.ApplicationWindow;
import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.widgets.*;

import java.io.File;

public class WordFrame extends ApplicationWindow {

    private OleFrame frame;
    private OleClientSite clientSite;
    private Shell shell;
    private Display display;



    public WordFrame() {
        super(null);
        this.display = new Display();
        this.shell = new Shell(display);
        this.shell.setText("Integration Word - Java");
        this.shell.setLayout(new FillLayout());
        try {
            this.frame = new OleFrame(shell, SWT.NONE);
            this.clientSite = new OleClientSite(this.frame, SWT.NONE, "Word.Document");
            this.clientSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
            addFileMenu();
        } catch (SWTError e) {
            e.printStackTrace();
            display.dispose();
        }
    }

    private void addFileMenu(){
        final Shell shell = this.frame.getShell();
        Menu menuBar = shell.getMenuBar();
        if(menuBar == null){
            menuBar = new Menu(shell, SWT.BAR);
            shell.setMenuBar(menuBar);
        }

        MenuItem fileMenu = new MenuItem(menuBar, SWT.CASCADE);
        fileMenu.setText("&Fichier");
        Menu menuFile = new Menu(fileMenu);
        fileMenu.setMenu(menuFile);
        frame.setFileMenus(new MenuItem[] { fileMenu });

        MenuItem menuFileOpen = new MenuItem(menuFile, SWT.CASCADE);
        menuFileOpen.setText("Ouverture...");
        menuFileOpen.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
                fileOpen();
            }
        });

        MenuItem menuFileSave = new MenuItem(menuFile, SWT.CASCADE);
        menuFileSave.setText("enregistrement...");
        menuFileSave.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
                fileSaveAs();
            }
        });

        MenuItem menuSpellCheck = new MenuItem(menuFile, SWT.CASCADE);
        menuSpellCheck.setText("SpellCheck");
        menuSpellCheck.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
                spellCheck();
            }
        });

        MenuItem menuFileExit = new MenuItem(menuFile, SWT.CASCADE);
        menuFileExit.setText("Quitter");
        menuFileExit.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
                shell.dispose();
            }
        });
    }

    private void fileOpen() {
        FileDialog dialog = new FileDialog(clientSite.getShell(), SWT.OPEN);
        dialog.setFilterExtensions(new String[] { "*.doc", "*.docx" });
        String fileName = dialog.open();
        if (fileName != null) {
            clientSite.dispose();
            clientSite = new OleClientSite(frame, SWT.NONE, "Word.Document", new File(fileName));
            clientSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
        }
    }

    /**
     *
     */
    private void fileSaveAs() {
        FileDialog dialog = new FileDialog(shell, SWT.SAVE);
        dialog.setFilterExtensions(new String[] { "*.doc", "*.docx" });
        String path = dialog.open();
        if (path != null) {
            if (clientSite.save(new File(path), false)) {
                MessageBox msgBox = new MessageBox(shell, SWT.ICON_INFORMATION);
                msgBox.setText("Sauvegarde");
                msgBox.setMessage("Sauvegarde rÃ©ussie");
                msgBox.open();
            }
            else
            {
                MessageBox msgBox = new MessageBox(shell, SWT.ICON_ERROR);
                msgBox.setText("Sauvegarde");
                msgBox.setMessage("Echec de la sauvegarde");
                msgBox.open();
            }
        }
    }

    /**
     *
     */
    private void spellCheck() {
        if ((clientSite.queryStatus(OLE.OLECMDID_SPELL) & OLE.OLECMDF_ENABLED) != 0) {
            clientSite.exec(OLE.OLECMDID_SPELL, OLE.OLECMDEXECOPT_PROMPTUSER, null, null);
        }
    }

    public void openning() {
        shell.open();

        while (!shell.isDisposed()) {
            if (!display.readAndDispatch())
                display.sleep();
        }
        display.dispose();
    }

}
