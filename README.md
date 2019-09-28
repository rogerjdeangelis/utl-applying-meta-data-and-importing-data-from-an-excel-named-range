# utl-applying-meta-data-and-importing-data-from-an-excel-named-range
Applying meta data and importing data from an excel named range

    Applying meta data and importing data from an excel named range                                                  
                                                                                                                     
    The op should create a named range for the retangle she/he is interested in.                                     
    Otherwise use 'sheet$A1:Z99', better to add named range.                                                         
                                                                                                                     
                                                                                                                     
    https://tinyurl.com/y2zoyno9                                                                                     
    https://github.com/rogerjdeangelis/utl-applying-meta-data-and-importing-data-from-an-excel-named-range           
                                                                                                                     
    SAS Forum                                                                                                        
    https://tinyurl.com/y3ughenu                                                                                     
    https://communities.sas.com/t5/SAS-Data-Management/proc-import-xlsx-secific-row/m-p/592343                       
                                                                                                                     
    *_                   _                                                                                           
    (_)_ __  _ __  _   _| |_                                                                                         
    | | '_ \| '_ \| | | | __|                                                                                        
    | | | | | |_) | |_| | |_                                                                                         
    |_|_| |_| .__/ \__,_|\__|                                                                                        
            |_|                                                                                                      
    ;                                                                                                                
                                                                                                                     
    %utlfkil(d:/xls/have.xlsx);                                                                                      
    libname xel "d:/xls/have.xlsx";                                                                                  
                                                                                                                     
    data xel.have;                                                                                                   
                                                                                                                     
      info="LABEL"; name="Student Name"; sex="Student Gender"; agec="Student Age";                                   
      output;                                                                                                        
                                                                                                                     
      do until (dne);                                                                                                
                                                                                                                     
         set sashelp.class(obs=3 keep= name sex age) end=dne;                                                        
                                                                                                                     
         info="DATA";                                                                                                
         agec=put(age,3.);                                                                                           
         output;                                                                                                     
                                                                                                                     
      end;                                                                                                           
                                                                                                                     
      drop age;                                                                                                      
      stop;                                                                                                          
    run;quit;                                                                                                        
                                                                                                                     
    libname xel clear;                                                                                               
                                                                                                                     
                                                                                                                     
     d:/xls/Have.xlsx                                                                                                
                                                                                                                     
       +---------------------------------------------------------+                                                   
       |  A        |  B           |  C            |  D           |                                                   
       +---------------------------------------------------------+                                                   
     1 |INFO       |NAME          |SEX            |AGEC          |                                                   
       +-----------+--------------+---------------+--------------|                                                   
     2 |LABEL      |Student Name  |Student Gender |Student AGe   |                                                   
       +-----------+--------------+---------------+--------------+                                                   
     3 |Alice      |    13        |      F        |    56.5      |                                                   
       +-----------+--------------+---------------+--------------+                                                   
     4 |Barbara    |    13        |      F        |    65.3      |                                                   
       +-----------+--------------+---------------+--------------+                                                   
     5 |Carol      |    14        |      F        |    62.8      |                                                   
       +-----------+--------------+---------------+--------------+                                                   
     6 |Henry      |    14        |      M        |    63.5      |                                                   
       -----------------------------------------------------------                                                   
       ...                                                                                                           
       [HAVE]                                                                                                        
                                                                                                                     
    *            _               _                                                                                   
      ___  _   _| |_ _ __  _   _| |_                                                                                 
     / _ \| | | | __| '_ \| | | | __|                                                                                
    | (_) | |_| | |_| |_) | |_| | |_                                                                                 
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                
                    |_|                                                                                              
    ;                                                                                                                
                                                                                                                     
                                                                                                                     
    WORK.WANT total obs=3                                                                                            
                                                                                                                     
    proc print data=want label;                                                                                      
    run;quit;                                                                                                        
                                                                                                                     
    INFO     Student Name    Student Gender    Student Age                                                           
                                                                                                                     
    DATA     Alfred          M                  14                                                                   
    DATA     Alice           F                  13                                                                   
    DATA     Barbara         F                  13                                                                   
                                                                                                                     
                                                                                                                     
    CONTENTS                                                                                                         
    ========                                                                                                         
                     Variables in Creation Order                                                                     
                                                                                                                     
     Variable    Type    Len    Format    Informat    Label                                                          
                                                                                                                     
     INFO        Char      5    $5.       $5.         INFO                                                           
     NAME        Char     12    $12.      $12.        Student Name                                                   
     SEX         Char     14    $14.      $14.        Student Gender                                                 
     AGEC        Char     11    $11.      $11.        Student Age                                                    
                                                                                                                     
                                                                                                                     
    *                                                                                                                
     _ __  _ __ ___   ___ ___  ___ ___                                                                               
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                              
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                              
    | .__/|_|  \___/ \___\___||___/___/                                                                              
    |_|                                                                                                              
    ;                                                                                                                
                                                                                                                     
                                                                                                                     
    libname xel "d:/xls/have.xlsx";                                                                                  
                                                                                                                     
    data want (drop=lbl);                                                                                            
                                                                                                                     
       set xel.have end=dne;;                                                                                        
       if _n_>1 then output want;                                                                                    
                                                                                                                     
       array chrs _character_;                                                                                       
                                                                                                                     
       length lbl $4096.;                                                                                            
       retain lbl;                                                                                                   
                                                                                                                     
       if _n_=1 then do;                                                                                             
          do over chrs;                                                                                              
             lbl=catx(" ",lbl, vname(chrs),"=",quote(strip(chrs)));                                                  
             call symputx("lbl",lbl);                                                                                
          end;                                                                                                       
       end;                                                                                                          
                                                                                                                     
    run;quit;                                                                                                        
                                                                                                                     
    /*                                                                                                               
    %put &=lbl;                                                                                                      
                                                                                                                     
    labels                                                                                                           
                                                                                                                     
    %put &=lbl;                                                                                                      
                                                                                                                     
    INFO = "LABEL"                                                                                                   
    NAME = "Student Name"                                                                                            
    SEX  = "Student Gender"                                                                                          
    AGEC = "Student Age"                                                                                             
    */                                                                                                               
                                                                                                                     
    proc datasets lib=work;                                                                                          
      modify want;                                                                                                   
      label &lbl;                                                                                                    
    run;quit;                                                                                                        
                                                                                                                     
    libname xel clear;                                                                                               
                                                                                                                     
