package cppMultiplyBy;

/**
 * @version 1.13 2019-01-20 
 * @author Paul I Ighofose
 *
 */
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.print.PrinterException;
import java.io.File;

import javax.swing.JColorChooser;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.colorchooser.AbstractColorChooserPanel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;

import org.eclipse.swt.SWT;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.sun.jna.Library;
import com.sun.jna.Native;
import com.sun.jna.ptr.DoubleByReference;

public class multiplyBy {
	
	public static void main(String[] args) {
		EventQueue.invokeLater(() -> {
			JFrame frame = new TableFrame();
			frame.setTitle("C++ multiplyBy Table Test");
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
			frame.setVisible(true);
		});
	}
}


/**
 * This frame contains an editable table. <br />
 */
class TableFrame extends JFrame {
	private static final long serialVersionUID = 1L;
	private String[] columnNames = {"Type", "Number1","Number2", "Multiplied Result"};
	private static String[][] format= {{"General", "%f"}, {"#,###,##0.00", "%,.2f"}, {"#,###,##0", "%,.0f"}};
	private static JComboBox<String> formats = new JComboBox<>();
	private String[] function = {"", "Print", "Send to New Email", "Send to Open Email", "Send to New Word Document", "Send to Open Word Document"};
	private JComboBox<String> functions = new JComboBox<>(function);
	private boolean formatting = false;
	private Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
	private static StringBuilder data;
	public static OleClientSite wdSite;
	
	public interface NativeMath extends Library {
		public void multiplyBy(DoubleByReference x, DoubleByReference y);
	}

	private NativeMath cpp = (NativeMath) Native.loadLibrary("C:\\Users\\Paul\\Documents\\Source Files\\dll\\multiplyBy x64.dll", NativeMath.class);		

	private Object[][] cells = {
			{". Total", 0, 0, 0},
			{"New", "", "", ""}
	};
	
	public TableFrame() {
		JTabbedPane tabbedPane = new JTabbedPane();
		DefaultTableModel model = new DefaultTableModel(cells, columnNames) {
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int col) {
				switch (col) {
				case 0:
					return false;
				case 1:
					return true;
				case 2:
					return true;
				case 3:
					return false;
				default:
					return true;
				}
			}
		};

		JColorChooser bgColourChooser = new JColorChooser();
		AbstractColorChooserPanel[] bgPanels = bgColourChooser.getChooserPanels();
		for (AbstractColorChooserPanel panel : bgPanels) {
			if (!panel.getDisplayName().equalsIgnoreCase("RGB")) {
				bgColourChooser.removeChooserPanel(panel);
			} else {
				bgColourChooser.getSelectionModel().setSelectedColor(new Color(0, 0, 108));
			}
		}

		JColorChooser fgColourChooser = new JColorChooser();
		AbstractColorChooserPanel[] fgPanels = fgColourChooser.getChooserPanels();
		for (AbstractColorChooserPanel panel : fgPanels) {
			if (!panel.getDisplayName().equalsIgnoreCase("RGB")) {
				fgColourChooser.removeChooserPanel(panel);
			} else {
				fgColourChooser.getSelectionModel().setSelectedColor(new Color(255, 255, 255));
			}
		}

		DefaultTableCellRenderer h = new DefaultTableCellRenderer(){
			private static final long serialVersionUID = 1L;

			@Override
			public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
					boolean hasFocus, int row, int column) {
				super.getTableCellRendererComponent(table, value, isSelected,
						hasFocus, row, column);
				setBackground(bgColourChooser.getSelectionModel().getSelectedColor());
				setForeground(fgColourChooser.getSelectionModel().getSelectedColor());
				return this;
			}
		};
		DefaultTableCellRenderer r = new DefaultTableCellRenderer(){
			private static final long serialVersionUID = 1L;

			@Override
			public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
					boolean hasFocus, int row, int column) {
				super.getTableCellRendererComponent(table, value, isSelected,
						hasFocus, row, column);
				setHorizontalAlignment(RIGHT);
				return this;
			}
		};
		JTable table = new JTable(model); 
		for (int i = 1; i <= 3; i++) {
			table.getColumnModel().getColumn(i).setCellRenderer(r);
		}
		table.getTableHeader().setDefaultRenderer(h);
		table.addKeyListener(new KeyListener() {
			
			@Override
			public void keyTyped(KeyEvent e) {
				// TODO Auto-generated method stub
			}
			
			@Override
			public void keyReleased(KeyEvent e) {
				// TODO Auto-generated method stub
				if (e.getKeyChar() == "".hashCode()) {
					JTable target = (JTable)e.getSource();
					data = new StringBuilder();
					data.append("Number1\tNumber2\tMultiplied Result\n");
					String totals = "\n";
					for (int i = 0; i < target.getRowCount(); i++) {
						if (target.getValueAt(i, 0) == "") {
							data.append(
										restyle(2, normVal(target.getValueAt(i, 1).toString())) + "\t" +
										restyle(2, normVal(target.getValueAt(i, 2).toString())) + "\t" +
										restyle(2, normVal(target.getValueAt(i, 3).toString())) + "\n"
									);
						} else if (target.getValueAt(i, 0) == ". Total") {
							totals +=  
									restyle(2, normVal(target.getValueAt(i, 1).toString())) + "\t" +
									restyle(2, normVal(target.getValueAt(i, 2).toString())) + "\t" +
									restyle(2, normVal(target.getValueAt(i, 3).toString())) + "\n";
						}
					}
					data.append(totals);
					StringSelection sel = new StringSelection(data.toString());
					clipboard.setContents(sel, null);		
				} else if (formatting == false) {
					JTable target = (JTable)e.getSource();
					double x = 0;
					double y = 0;
					double z = 0;
					int t = 0;
					int n = 0;
					for (int i = 0; i < target.getRowCount(); i++) {
						if (target.getValueAt(i, 0) == ". Total") {
							t = i;
						}
						if (target.getValueAt(i, 0) == "New") {
							n = i;
						}
					}
					if (target.getSelectedColumn() != 3 && target.getSelectedColumn() != 0) {
						int row = target.getSelectedRow();
						if (target.getValueAt(row, 0) == "New") {
							target.setValueAt("", row, 0);
							model.addRow(new Object[] {"New", "", "", ""});
							if (t == row - 1 || (!(t <= 1) && !(t >= target.getRowCount() - 2))) {
								model.moveRow(t, t, row);
							} else if (t == 1 && n == 0) {
								model.moveRow(row, row, target.getRowCount() - 1);
								model.moveRow(target.getRowCount() - 1, target.getRowCount() - 1, 0);
							} 
						} else {
							x = normVal(target.getValueAt(row, 1).toString());
							y = normVal(target.getValueAt(row, 2).toString());
							z = normVal(calculate(x, y));
							target.setValueAt(z, row, 3);
							if (target.getRowCount() > 3) {
								x = 0;
								y = 0;
								z = 0;
								for (int i = 0; i < target.getRowCount(); i++) {
									if (target.getValueAt(i, 0) == "") {
										x += normVal(target.getValueAt(i, 1).toString());
										y += normVal(target.getValueAt(i, 2).toString());
										z += normVal(target.getValueAt(i, 3).toString());
									}
								}
							}
							target.setValueAt(restyle(2, x), t, 1);
							target.setValueAt(restyle(2, y), t, 2);
							target.setValueAt(restyle(2, z), t, 3);
						}
					}
				}
			}
			
			@Override
			public void keyPressed(KeyEvent e) {
				// TODO Auto-generated method stub
			}
		});
		table.setAutoCreateRowSorter(true);
		JScrollPane tblPane = new JScrollPane(table);
		tblPane.setPreferredSize(new Dimension(700, 550));
		tabbedPane.addTab("Table", tblPane);
		tabbedPane.addTab("Header Background Colour", bgColourChooser);
		tabbedPane.addTab("Header Fore Colour", fgColourChooser);
		add(tabbedPane, BorderLayout.NORTH);

		functions.addActionListener(event -> {
				if (functions.getItemAt(functions.getSelectedIndex()) != "") {
					if (functions.getItemAt(functions.getSelectedIndex()) == "Print") {
						try {table.print(); }
						catch (SecurityException | PrinterException ex) { ex.printStackTrace(); }
					} else if (functions.getItemAt(functions.getSelectedIndex()) == "Send to New Email") {
						String tbl = TableLayouts.tblHTML(table, bgColourChooser.getSelectionModel().getSelectedColor(), fgColourChooser.getSelectionModel().getSelectedColor());
						OutlookAutomation.sendToEmail("C++ multiplyBy Java Test", tbl, "New");
					} else if (functions.getItemAt(functions.getSelectedIndex()) == "Send to Open Email") {
						String tbl = TableLayouts.tblHTML(table, bgColourChooser.getSelectionModel().getSelectedColor(), fgColourChooser.getSelectionModel().getSelectedColor());
						OutlookAutomation.sendToEmail("C++ multiplyBy Java Test", tbl, "Open");
						table.grabFocus();
						table.requestFocus();
					} else if (functions.getItemAt(functions.getSelectedIndex()) == "Send to New Word Document") {
						WordAutomation.sendToWord("New", table, bgColourChooser.getSelectionModel().getSelectedColor(), fgColourChooser.getSelectionModel().getSelectedColor());
					} else if (functions.getItemAt(functions.getSelectedIndex()) == "Send to Open Word Document") {
						WordAutomation.sendToWord("Open", table, bgColourChooser.getSelectionModel().getSelectedColor(), fgColourChooser.getSelectionModel().getSelectedColor());
					}
					
					functions.setSelectedIndex(0);
				}
		});

		for (int i = 0; i < format.length; i++) {
			formats.addItem(format[i][0]);
		}
		formats.addActionListener(event -> {
			formatting = true;
			for (int i = 0; i < table.getRowCount(); i++) {
				if (table.getModel().getValueAt(i, 0) == null || table.getModel().getValueAt(i, 0) == "") {
					for (int j = 1; j <= 3; j++) {
						table.getModel().setValueAt(restyle(1, normVal(table.getModel().getValueAt(i, j).toString())), i, j);
					}
				} else if (table.getModel().getValueAt(i, 0) == ". Total") {
					for (int j = 1; j <= 3; j++) {
						table.getModel().setValueAt(restyle(2, normVal(table.getModel().getValueAt(i, j).toString())), i, j);
					}
				}
			}
			formatting = false;
		});
		
		JPanel buttonPanel = new JPanel();
		buttonPanel.add(formats);
		buttonPanel.add(functions);
		
		add(buttonPanel, BorderLayout.SOUTH);
		table.setDragEnabled(true);
		bgColourChooser.setDragEnabled(true);
		fgColourChooser.setDragEnabled(true);
		setSize(700, 580);
		pack();
	}
	
	private String calculate(double xin, double yin) {
		DoubleByReference x = new DoubleByReference();
		x.setValue(xin);;
		DoubleByReference y = new DoubleByReference();
		y.setValue(yin);
				
		cpp.multiplyBy(x , y);
		return String.format("%f", y.getValue());
	}
	
	public static double normVal(String x) {
		String out = x.replaceAll(",", "").replaceAll("'", "");
		if (out == "") {
			out = "0";
		}
		return Double.parseDouble(out);
	}
	
	public static String restyle(int typ, double val) {
		if (formats.getSelectedIndex() == 0) {
			return String.format("%f", val);
		} else if (typ == 1) {
			return String.format("%,f", val);
		} else {
			return String.format(format[formats.getSelectedIndex()][1], val);
		}
	}
	
}

class OutlookAutomation {
	private static String[] attachmentPaths = null;

	public OutlookAutomation() {
		
	}
	
	public static void sendToEmail(String sbj, String htm, String type) {
		Display  display = Display.getCurrent();
		Shell shell = new Shell(display);
		OleFrame frame = new OleFrame(shell, SWT.NONE);
		OleClientSite site = new OleClientSite(frame, SWT.NONE, "OVCtl.OVCtl");
		site.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
		OleClientSite site2 = new OleClientSite(frame, SWT.NONE, "Outlook.Application");
		OleAutomation outlook = new OleAutomation(site2);
		OleAutomation oMail = null;
		if (type == "New") {
			oMail = invoke(outlook, "CreateItem", 0).getAutomation();
			invoke(oMail, "Display");
			setProperty(oMail, "BodyFormat", 2);
			setProperty(oMail, "Subject", sbj);
			setProperty(oMail, "HtmlBody", htm);
			
			if (null != attachmentPaths) {
				for (String attachmentPath : attachmentPaths) {
					File file = new File(attachmentPath);
					if (file.exists()) {
						OleAutomation attachments = getProperty(oMail, "Attachments");
						invoke(attachments, "Add", attachmentPath);
					}
				}
			}				
		} else {
			OleAutomation inspector = invoke(outlook, "ActiveInspector").getAutomation();
			OleAutomation editor = invoke(inspector, "WordEditor").getAutomation();
			OleAutomation app = editor.getProperty(1).getAutomation();
			setProperty(app, "Selection", "placeHere");
			oMail = invoke(inspector, "CurrentItem").getAutomation();
			setProperty(oMail, "HtmlBody", getProperty(oMail, "HtmlBody", "").replace("placeHere", htm));
		}
	}	
	
	protected static OleAutomation getProperty(OleAutomation auto, String name) {
		Variant varResult = auto.getProperty(property(auto, name));
		if (varResult != null && varResult.getType() != OLE.VT_EMPTY) {
			OleAutomation result = varResult.getAutomation();
			varResult.dispose();
			return result;
		}
		return null;
	}
	
	private static String getProperty(OleAutomation auto, String name, String n) {
		Variant varResult = auto.getProperty(property(auto, name));
		if (varResult != null && varResult.getType() != OLE.VT_EMPTY) {
			String result = varResult.getString();
			varResult.dispose();
			return result;
		}
		return null;
	}
	
	protected static Variant invoke(OleAutomation auto, String command, String value) {
		return auto.invoke(property(auto, command), new Variant[] { new Variant(value)});
	}
	
	protected static Variant invoke(OleAutomation auto, String command) {
		return auto.invoke(property(auto, command));
	}
	
	private static Variant invoke(OleAutomation auto, String command, int value) {
		return auto.invoke(property(auto, command), new Variant[] { new Variant(value)});
	}
	
	protected static boolean setProperty(OleAutomation auto, String name, String value) {
		return auto.setProperty(property(auto, name), new Variant(value));
	}

	protected static boolean setProperty(OleAutomation auto, String name, int value) {
		return auto.setProperty(property(auto, name), new Variant(value));
	}
	
	protected static int property(OleAutomation auto, String name) {
		return auto.getIDsOfNames(new String[] {name})[0];
	}
}

class WordAutomation extends OutlookAutomation {
	public WordAutomation() {
		
	}

	public static void sendToWord(String type, JTable table, Color bg, Color fg) {
		long bgColour = (long) (bg.getRed() + (bg.getGreen() * 256) + (bg.getBlue() * 65536));
		long fgColour = (long) (fg.getRed() + (fg.getGreen() * 256) + (fg.getBlue() * 65536));
		Display  display = Display.getCurrent();
		Shell shell = new Shell(display);
		OleFrame frame = new OleFrame(shell, SWT.NONE);
		if (TableFrame.wdSite == null) {
			TableFrame.wdSite = new OleClientSite(frame, SWT.NONE, "Word.Application");
			TableFrame.wdSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
			type = "New";
		}
		OleAutomation word = new OleAutomation(TableFrame.wdSite);
		OleAutomation application;
		try {
			application = getProperty(word, "Application");
		} catch (Exception e) {
			TableFrame.wdSite = new OleClientSite(frame, SWT.NONE, "Word.Application");
			TableFrame.wdSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
			word = new OleAutomation(TableFrame.wdSite);
			application = getProperty(word, "Application");
			type = "New";			
		}
		setProperty(application, "Visible", 1);
		OleAutomation documents = getProperty(application, "Documents");
		if (type == "New") {
			invoke(documents, "Add").getAutomation();
		}
		OleAutomation selection =  getProperty(application, "Selection");
		setProperty(application, "Selection", "placeHere");
		invoke(selection, "ConvertToTable");
		OleAutomation borders = getProperty(selection, "Borders"); 
		setProperty(borders, "InsideLineStyle", 1);
		setProperty(borders, "OutsideLineStyle", 1);
		OleAutomation cells = getProperty(selection, "Cells");
		Variant[] rgvarg = new Variant[2];
		rgvarg[0] = new Variant((int) table.getRowCount() + 1);
		rgvarg[1] = new Variant((int) table.getColumnCount() - 1);
		invoke(cells, "Split", rgvarg);
		borders = getProperty(cells, "Borders");
		setProperty(borders, "InsideLineStyle", 1);
		OleAutomation colour;
		OleAutomation font;
		OleAutomation align;
		for (int col = 1; col < table.getColumnCount(); col++) {
			invoke(selection,"SelectCell");
			setProperty(selection,"Text", table.getColumnName(col));
			colour = getProperty(selection, new String[] {"Range", "Shading"});
			font = getProperty(selection, new String[] {"Range", "Font"});
			setProperty(colour, "BackgroundPatternColor", bgColour);
			setProperty(font, "Color", fgColour);
			setProperty(font, "Bold", 1);
			if (col < table.getColumnCount() - 1) {
				invoke(selection,"MoveRight");
			}
		}
		for (int row = 0; row < table.getRowCount(); row++) {
			if (table.getValueAt(row, 0) == "") {
				invoke(selection,"MoveDown");
				for (int col = 1; col < table.getColumnCount(); col++) {
					if (col < table.getColumnCount() - 1) {
						invoke(selection,"MoveLeft");
					}
				}
				for (int col = 1; col < table.getColumnCount(); col++) {
					invoke(selection,"SelectCell");
					setProperty(selection,"Text", TableFrame.restyle(2, TableFrame.normVal(table.getValueAt(row, col).toString())));
					align = getProperty(selection, new String[] {"Range", "ParagraphFormat"});
					setProperty(align, "Alignment", 2);
					if (col < table.getColumnCount() - 1) {
						invoke(selection,"MoveRight");
					}
				}
			}
		}
		invoke(selection,"MoveDown");
		for (int col = 1; col < table.getColumnCount(); col++) {
			setProperty(borders, "OutsideLineStyle", 0);
			if (col < table.getColumnCount() - 1) {
				borders = getProperty(cells, "Borders");
				setProperty(borders, "InsideLineStyle", 0);
				invoke(selection,"MoveLeft");
			}
			setProperty(borders, "OutsideLineStyle", 0);
		}
		invoke(selection,"MoveUp");
		invoke(selection,"SelectRow");
		setProperty(borders, "OutsideLineStyle", 1);
		invoke(selection,"MoveDown");
		invoke(selection,"MoveDown");
		invoke(selection,"SelectRow");
		setProperty(borders, "OutsideLineStyle", 1);
		invoke(selection,"MoveUp");
		invoke(selection,"SelectCell");
		for (int row = 0; row < table.getRowCount(); row++) {
			if (table.getValueAt(row, 0) == ". Total") {
				invoke(selection,"MoveDown");
				for (int col = 1; col < table.getColumnCount(); col++) {
					invoke(selection,"SelectCell");
					setProperty(selection,"Text", TableFrame.restyle(2, TableFrame.normVal(table.getValueAt(row, col).toString())));
					align = getProperty(selection, new String[] {"Range", "ParagraphFormat"});
					setProperty(align, "Alignment", 2);
					if (col < table.getColumnCount() - 1) {
						invoke(selection,"MoveRight");
					}
				}
			}
		}
	}
	
	private static Variant invoke(OleAutomation auto, String command, Variant[] value) {
		return auto.invoke(property(auto, command), value);
	}
	
	private static OleAutomation getProperty(OleAutomation auto, String[] names) {
		OleAutomation newAuto = auto;
		for (int i = 0; i < names.length; i++) {
			newAuto = getProperty(newAuto, names[i]); 
		}
		return newAuto;
	}

	protected static boolean setProperty(OleAutomation auto, String name, long value) {
		return auto.setProperty(property(auto, name), new Variant(value));
	}
}

class TableLayouts {
	
	public TableLayouts() {
		
	}
	
	public static String tblHTML(JTable table, Color bgColour, Color fgColour) {
		JTable target = table;
		StringBuilder stl = new StringBuilder(); 
		StringBuilder tbl = new StringBuilder(); 
		StringBuilder tr = new StringBuilder(); 
		StringBuilder th = new StringBuilder(); 
		StringBuilder td = new StringBuilder(); 
		StringBuilder tt = new StringBuilder(); 
	    stl.append("<style>\n");
	    stl.append("table {\n");
	    stl.append("\tborder-collapse: collapse;\n");
		stl.append("}\n");
	    stl.append("table, th, td {\n");
	    stl.append("\tpadding: 0px 5px 0px 5px;\n");
	    stl.append("\tborder: 1 solid black;\n");
	    stl.append("\tfont-size: 1em;\n");
		stl.append("}\n");
	    stl.append("td {\n");
		stl.append("\tcolor: black;\n");
	    stl.append("\ttext-align: right;\n");
		stl.append("}\n");
	    stl.append("th {\n");
		stl.append("\tcolor: rgb(" + fgColour.getRed() + ", " + fgColour.getGreen() + ", " + fgColour.getBlue() + ");\n");
		stl.append("\tbackground-color: rgb(" + bgColour.getRed() + ", " + bgColour.getGreen() + ", " + bgColour.getBlue() + ");\n");
		stl.append("}\n");
	    stl.append("</style>\n");
		tbl.append("<table id=\"mulitplyBy\">\n");
	    String trd = "\t<tr>\n";
	    String tre = "\n\t</tr>\n";
	    String thd = "\t\t<th>\n\t\t\t";
	    String the = "\n\t\t</th>\n";
	    String tdd = thd.replace("th", "td");
	    String tde = the.replace("th", "td");
	    for (int col = 1; col < target.getColumnCount(); col++) {
	        th.append(thd + target.getColumnName(col) + the);
	    }
	    tbl.append(th);
	    for (int row = 0; row < target.getRowCount(); row++) {
            td = new StringBuilder();
            for (int col = 1; col < target.getColumnCount(); col++) {
            	if (target.getValueAt(row, 0) == ". Total") {
                    tt.append(tdd + TableFrame.restyle(2, TableFrame.normVal(target.getValueAt(row, col).toString())) + tde);
                } else if (target.getValueAt(row, 0) == "") {
                    td.append(tdd + TableFrame.restyle(2, TableFrame.normVal(target.getValueAt(row, col).toString())) + tde);
                }
            }
            if (target.getValueAt(row, 0) == "") {
            	tr.append(trd + td + tre);
            } 
	    }
	    tr.append(trd + "<td colspan=\"3\" style=\"border-left: none; border-right: none;\">  &nbsp; </td>" + tre);	    
	    tr.append(trd + tt + tre);
	    tbl.append(tr + "</table>");
	    stl.append(tbl);
		return stl.toString();
	}
}