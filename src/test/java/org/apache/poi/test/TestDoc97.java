package org.apache.poi.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
public class TestDoc97 {
	   private static ArrayList<RawTable> m_RawTables =new ArrayList<RawTable>();
	   public static String docTableCellHyperlinkText(String cellText) {
			return cellText.replaceAll("HYPERLINK \"mailto:[^@]+@{1}[^@^\"]+\"{1}",
					"");
		}
		public static String formatToSpace(String s) {
			byte[] bs = s.getBytes();
			for (int i = 0; i < bs.length; ++i) {
				if (bs[i] < 32 && bs[i] > 0)
					bs[i] = 32;
			}
			return new String(bs);
		}
		 public static String byteHEX(byte ib) {
	          char[] Digit = { '0','1','2','3','4','5','6','7','8','9',
	          'A','B','C','D','E','F' };
	          char [] ob = new char[2];
	          ob[0] = Digit[(ib >>> 4) & 0X0F];
	          ob[1] = Digit[ib & 0X0F];
	          String s = new String(ob);
	          return s;
	    }
	  public static void main(String[] args) {
			POIFSFileSystem pfs;
//			byte[]parray=new byte[64];			 
//	    	 MessageDigest temp=null;
//	    	 try {
//			    temp = MessageDigest.getInstance("MD5");
//			 } catch (NoSuchAlgorithmException e) {
//			 }
//	    	// System.out.println("hhh"+temp.digest()[0]);
//	    	 temp.update("123456".getBytes());
//	    	 byte []yy=temp.digest();
//	    	 for(int i=0;i<yy.length;i++)
//	    	 {
//	    		 //System.out.print(byteHEX(yy[i]));
//	    	 }
			HWPFDocument hwpf =null;
			try {
				pfs = new POIFSFileSystem(new FileInputStream("./test/20030523jm.doc"));
			    hwpf = new HWPFDocument(pfs,"111111"); 
			} catch ( Exception e) {
			}		
			Range range = hwpf.getOverallRange();
			TableIterator it = new TableIterator(range);
			while (it.hasNext()) {
				RawTable rawTable = new RawTable();
				Table tb = (Table) it.next();
				for (int i = 0; i < tb.numRows(); i++) {
					ArrayList<CellShap> r = new ArrayList<CellShap>();
					TableRow tr = tb.getRow(i);
					for (int j = 0; j < tr.numCells(); j++) {
						TableCell tc = tr.getCell(j);
						StringBuilder sb = new StringBuilder();
						for (int k = 0; k < tc.numParagraphs(); k++) {
							Paragraph para = tc.getParagraph(k);
							sb.append(docTableCellHyperlinkText(
									formatToSpace(para.text())).trim()
									+ " ");
							System.out.println(sb.toString());
						} 
						r.add(new CellShap(sb.toString().trim(), tc.getLeftEdge(), -1, -1));
					} 
					rawTable.add(r);
					System.out.println("");
				} 
				m_RawTables.add(rawTable);
			} 
			}
	  }

