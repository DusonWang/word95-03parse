package org.apache.poi.hwpf;

import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

public class HWPFRC4 {
	 public byte[]state;
	 public int x;
	 public int y;
     public HWPFRC4(){ 
    	 state=new byte[256];
     }
     public void preparekey(byte[] key_data_ptr, int key_data_len, HWPFRC4 key)
     {
    	    int index1;
    	    int index2;
    	    byte []state=new byte[256];
    	    int counter;
    	    state = key.state;
    	    for (counter = 0; counter < 256; counter++) state[counter] =  (byte)counter;
    	    key.x = 0;
    	    key.y = 0;
    	    index1 = 0; 
    	    index2 = 0;
    	    for (counter = 0; counter < 256; counter++)
    	      {
    		  index2 =  ((key_data_ptr[index1]&0xff) + (state[counter]&0xff )+ index2 ) &0xff;
    		  byte btemp=state[counter];
    		  state[counter]=state[index2];
    		  state[index2]=btemp;
    		  index1 = ((index1 + 1) % key_data_len);
    	      }
     }
     public void makekey(int block,HWPFRC4 rc4key,HWMPFEMD5 md5) 
     {   
	    	 byte[]pwarray=new byte[64];
	    	 HWMPFEMD5 temp=new HWMPFEMD5();
	    	 for(int i=0;i<64;i++) pwarray[i]=0;
	    	 for(int i=0;i<5;i++)
	    	 {  
	    		 pwarray[i]=md5.digest[i]; 
	    	 }  
	    	 pwarray[5] = (byte) (block & 0xFF);
	    	 pwarray[6] = (byte) ((block >> 8) & 0xFF);
	    	 pwarray[7] = (byte) ((block >> 16) & 0xFF);
	    	 pwarray[8] = (byte) ((block >> 24) & 0xFF);
	    	 pwarray[9] = (byte)0x80;
	    	 pwarray[56] =(byte)0x48;
	    	 temp.md5Init();
	    	 temp.md5Update(pwarray, 64);
	    	 temp.getMD5StoreDigest(temp);
	    	 preparekey(temp.digest, 16, rc4key);
     }
     void rc4 ( byte[] buffer_ptr, int buffer_len, HWPFRC4  key)
     {
	    	 int x;
	         int y;
	         byte []state=new byte[256];
	         int xorIndex;
	         int counter;
	         x = key.x;
	         y = key.y;
	         state = key.state;
	         for (counter = 0; counter < buffer_len; counter++)
	           {
	        	  x =  ((x + 1) & 0xff);
	        	  y = (((state[x]&0xff) + y) & 0xff);
	        	  byte btemp=state[x];
	    		  state[x]=state[y];
	    		  state[y]=btemp;
	        	  xorIndex =  (((state[x]&0xff )+ (state[y]&0xff)) & 0xff);
	        	  buffer_ptr[counter]^=(state[xorIndex]);
	           }
	         key.x = x; 
	         key.y = y;
     }
}
