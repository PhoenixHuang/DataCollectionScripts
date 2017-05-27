
' The input metadata document (.mdd) file
#define INPUTMDM "R:\share\TEST\old.mdd"

' The output metadata document (.mdd) file
#define OUTPUTMDM "R:\share\TEST\New.mdd"

' Copy the museum.mdd sample file so that
' we do not update the original file...
Dim fso, f 
Set fso = CreateObject("Scripting.FileSystemObject")  
fso.CopyFile(INPUTMDM, OUTPUTMDM, True)

dim f1, f2, f3, f4, f5,f6
set f1=fso.CreateTextFile("DataMap.sql")
set f2=fso.CreateTextFile("CodeMap.sql")      
set f3=fso.CreateTextFile("ExcelMap.csv")      
set f4=fso.CreateTextFile("Transform.sql")      
set f5=fso.CreateTextFile("FinalDataTable.sql")    
set f6=fso.CreateTextFile("QlikDataLoad.sql")  
' Make sure that the read-only attribute is not set
Set f = fso.GetFile(OUTPUTMDM)
If f.Attributes.BitAnd(1) Then
    f.Attributes = f.Attributes - 1
End If

Dim MDM2

' Create the MDM object
Set MDM2 = CreateObject("MDM.Document")
MDM2.Open (OUTPUTMDM) 


dim m,n
n=1
f3.writeline("Respondent.Serial,Long,Serial,int")

f5.writeline("CREATE Function [dbo].[Split] ")
f5.writeline("(@Sql nvarchar(4000), @Splits nvarchar(10)=',') returns @temp Table (Ans nvarchar(100)) ")
f5.writeline("As ")
f5.writeline("Begin ")
f5.writeline("Declare @i Int ")
f5.writeline("Set @Sql = RTrim(LTrim(@Sql)) ")
f5.writeline("Set @i = CharIndex(@Splits,@Sql) ")
f5.writeline("While @i >= 1 ")
f5.writeline("Begin ")
f5.writeline("Insert @temp Values(Left(@Sql,@i-1)) ")
f5.writeline("Set @Sql = SubString(@Sql,@i+1,Len(@Sql)-@i) ")
f5.writeline("Set @i = CharIndex(@Splits,@Sql) ")
f5.writeline("End ")
f5.writeline("If @Sql <> '' ")
f5.writeline("Insert @temp Values (@Sql) ")
f5.writeline("Return ")
f5.writeline("End ")
f5.writeline("GO")
f5.writeline("")



f5.writeline("Create Table FinalData (Serial int NOT NULL PRIMARY KEY,")
for each m in MDM2.Fields
    'GetDataMap(m,f1,f2)
    'debug.Log(m.name)
    GetExcelMap(m,f3,f5,"")
    GetSQL(m,n,f4)
    n=n+1    
next
f1.Close()
f2.Close()
f3.Close()
f4.Close()
f5.writeline(") --remove the last comma")
f5.writeline("")
f5.writeline("--bulk insert finaldata from 'R:\share\TEST\L.csv'") 
f5.writeline("--with(   ")
f5.writeline("--		FIELDTERMINATOR = ',', ")
f5.writeline("--        ROWTERMINATOR = '\n', ")
f5.writeline("--        FIRSTROW=2,")
f5.writeline("--        datafiletype='widechar'")
f5.writeline("--)")
f5.writeline("")
f5.Close()
f6.WriteLine("use BI;")
f6.WriteLine("print 'OLEDB CONNECT TO [Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BI;Data Source=.\sqlexpress;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=TSWHNL0034;Use Encryption for Data=False;Tag with column collation when possible=False];'")
f6.WriteLine("select 'SQL SELECT * FROM '+ TABLE_CATALOG+'.'+TABLE_SCHEMA+'.'+table_name+';' from INFORMATION_SCHEMA.tables")

f6.Close()
mdm2.Close()


sub GetDataMap(x,f1,f2)
    dim y,z,vType
    if not x.issystem then        
        select case x.ObjectTypeValue
            case ObjectTypesConstants.mtVariable
                vtype=getvartype(x)
                'debug.Log(x.name+";"+ctext(x.leveldepth))
                f1.writeline("insert into DataMap values('"+x.name+"','"+vType+"','"+replace(x.label,"'","")+"')")
                for each y in x.categories
                    f2.writeline("insert into CodeMap values('"+x.name+"','"+y.name +"','"+ replace(y.label,"'","")+"')")
                next
            case objectTypesconstants.mtarray 'loop    
                vType="Loop"
                'debug.Log(x.name+";"+ctext(x.leveldepth))
                f1.writeline("insert into DataMap values('"+x.name+"','"+vType+"','"+replace(x.label,"'","")+"')")                
                for each z in x.fields                    
                        GetDataMap(z,f1,f2)                        
                next                
                for each y in x.categories
                        f2.writeline("insert into CodeMap values('"+x.name+"','"+y.name +"','"+ replace(y.label,"'","")+"')")                            
                next            
                f1.writeline("--"+x.name+":End")
        end select    
        
    end if
 
end sub

sub GetExcelMap(x,f3,f5,m)
    
    dim y,z    
    select case x.ObjectTypeValue
            case ObjectTypesConstants.mtVariable                
                if x.leveldepth=1 then 
                    f3.writeline(x.name+","+GetVarType(x)+","+x.name+","+GetVarSQLType(x))
                    f5.writeline("["+x.name+"] "+GetVarSQLType(x)+",")
                end if    
            case objectTypesconstants.mtarray 'loop    
                m=m+x.name+"[{"
                dim a,b
                for each z in x.fields                    
                    for each y in x.categories
                        a=y.name+"}]"
                        if z.objecttypevalue=ObjectTypesConstants.mtVariable then                               
                                b=m+a+"."+z.name
                                f3.writeline(m+a+"."+z.name+","+GetVarType(z)+","+GetName(b)+","+GetVarSQLType(z))
                                f5.writeline("["+GetName(b)+"] "+GetVarSQLType(z)+",")
                        end if
                        GetExcelMap(z,f3,f5,m+a+".")  'great
                    next                    
                next                
            case objectTypesconstants.mtClass
'                m=m+x.name+"_"
'                for each z in x.fields
'                    if z.leveldepth>1 and z.objecttypevalue=ObjectTypesConstants.mtVariable then                               
'                        f3.writeline(m+z.name+","+GetVarType(z)+","+m+z.name)                
'                    end if
'                    GetExcelMap(z,f3,f5,m+"_")  'great
'                next    
    end select    
    
end sub

function GetVarType(x) 
    dim vType
    vType="Unknown"
    if x.ObjectTypeValue = ObjectTypesConstants.mtVariable then
        select case x.dataType
            case DataTypeConstants.mtText
                    vType="Text"
            case DataTypeConstants.mtLong
                    vType="Long"
            case DataTypeConstants.mtDouble
                    vType="Double"
            case DataTypeConstants.mtDate
                    vType="Date"
            case DataTypeConstants.mtBoolean
                    vType="Bool"
            case DataTypeConstants.mtCategorical
                 if x.maxvalue is null or x.maxvalue>1 then
                     vType="Multi"
                 else
                     vType="Single"
                 end if                    
        end select    
    end if
    GetVarType=vType
end function

function GetVarSQLType(x) 
    dim vType
    vType="nvarchar(500)"
    if x.ObjectTypeValue = ObjectTypesConstants.mtVariable then
        select case x.dataType
            case DataTypeConstants.mtText
                    vType="nvarchar(500)"
            case DataTypeConstants.mtLong
                    vType="int"
            case DataTypeConstants.mtDouble
                    vType="float"
            case DataTypeConstants.mtDate
                    vType="nvarchar(50)"
            case DataTypeConstants.mtBoolean
                    vType="nvarchar(10)"
            case DataTypeConstants.mtCategorical                 
                    vType="nvarchar(500)"                 
        end select    
    end if
    GetVarSQLType=vType
end function

'the DSL does not support recurssion...
Sub GetSql(x,cnt,f4)
    dim a,b,c,d
    dim y1,y2,y3,y4    
    if x.ObjectTypeValue=objectTypesconstants.mtarray then
        f4.writeline(GetDim(x))
        for each a in x.fields                    
            if a.ObjectTypeValue=objectTypesconstants.mtarray then
                f4.writeline(GetDim(a))
                for each b in a.fields
                    if b.ObjectTypeValue=objectTypesconstants.mtarray then
                        f4.writeline(GetDim(b))
                        for each c in b.fields
                            if c.ObjectTypeValue=objectTypesconstants.mtarray then
                                f4.writeline(GetDim(c))
                                for each d in c.fields
                                    if d.ObjectTypeValue=objectTypesconstants.mtarray then
                                        debug.MsgBox("5 level loops, please change the code.")
                                    else 'level 4 d
                                        f4.writeline(GetDim(d))
                                        if GetVarType(d)="Multi" then
                                            f4.writeline("--"+d.name)
                                            f4.writeline(";with T"+Ctext(cnt)+" as (Select [Serial],["+x.name+"],["+a.name+"],["+b.name+"],["+c.name+"],Codes from finaldata cross apply(values")                                        
                                            for each y1 in x.categories
                                                for each y2 in a.categories
                                                    for each y3 in b.categories
                                                        for each y4 in c.categories
                                                            f4.writeline("('"+Lcase(y1.name)+"','"+Lcase(y2.name)+"','"+Lcase(y3.name)+"','"+Lcase(y4.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"{"+y2.name+"}_"+b.name+"{"+y3.name+"}_"+c.name+"{"+y4.name+"}_"+d.name+"]),")
                                                        next
                                                    next
                                                next    
                                            next                            
                                            f4.writeline(") x("+x.name+","+a.name+","+b.name+","+c.name+",Codes) where Len(Codes)>0")
                                            f4.writeline(")select Serial,"+x.name+","+a.name+","+b.name+","+c.name+",Ans as "+d.name+"Ans into Fact"+d.name+" from T"+Ctext(cnt)+" cross apply split(Codes,';')")
                                            f4.writeline("ALTER TABLE Fact"+d.name+" ALTER Column "+d.name+"Ans NVARCHAR(500)")
                                            f4.writeline("ALTER TABLE Fact"+d.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+d.name+" FOREIGN KEY ("+d.name+"Ans) REFERENCES Dim"+d.name+"("+d.name+"Ans)")
                                            
                                        else
                                            f4.writeline("--"+d.name)
                                            f4.writeline("Select [Serial],"+x.name+","+a.name+","+b.name+","+c.name+","+d.name+" into Fact"+d.name+" from finaldata cross apply(values")                                        
                                            for each y1 in x.categories
                                                for each y2 in a.categories
                                                    for each y3 in b.categories
                                                        for each y4 in c.categories
                                                            f4.writeline("('"+Lcase(y1.name)+"','"+Lcase(y2.name)+"','"+Lcase(y3.name)+"','"+Lcase(y4.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"{"+y2.name+"}_"+b.name+"{"+y3.name+"}_"+c.name+"{"+y4.name+"}_"+d.name+"]),")
                                                        next
                                                    next
                                                next    
                                            next                            
                                            f4.writeline(") x("+x.name+","+a.name+","+b.name+","+c.name+","+d.name+") where Len("+d.name+")>0")
                                            if GetVarType(d)="Single" then f4.writeline("ALTER TABLE Fact"+d.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+d.name+" FOREIGN KEY ("+d.name+") REFERENCES Dim"+d.name+"("+d.name+")")
                                        end if
                                        f4.writeline("ALTER TABLE Fact"+d.name+" WITH NOCHECK ADD CONSTRAINT FK_Serial_"+d.name+" FOREIGN KEY (Serial) REFERENCES FinalData(Serial)")
                                        
                                        f4.writeline("ALTER TABLE Fact"+d.name+" ALTER Column "+c.name+" NVARCHAR(500)")
                                        f4.writeline("ALTER TABLE Fact"+d.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+" FOREIGN KEY ("+c.name+") REFERENCES Dim"+c.name+"("+c.name+")")
                                        f4.writeline("ALTER TABLE Fact"+d.name+" ALTER Column "+b.name+" NVARCHAR(500)")
                                        f4.writeline("ALTER TABLE Fact"+d.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+"_"+b.name+" FOREIGN KEY ("+b.name+") REFERENCES Dim"+b.name+"("+b.name+")")
                                        f4.writeline("ALTER TABLE Fact"+d.name+" ALTER Column "+a.name+" NVARCHAR(500)")
                                        f4.writeline("ALTER TABLE Fact"+d.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+"_"+b.name+"_"+c.name+" FOREIGN KEY ("+a.name+") REFERENCES Dim"+a.name+"("+a.name+")")
                                        f4.writeline("ALTER TABLE Fact"+d.name+" ALTER Column "+x.name+" NVARCHAR(500)")
                                        f4.writeline("ALTER TABLE Fact"+d.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+"_"+b.name+"_"+c.name+"_"+d.name+" FOREIGN KEY ("+x.name+") REFERENCES Dim"+x.name+"("+x.name+")")
                                    end if
                                next
                            else 'level 3 c
                                f4.writeline(GetDim(c))
                                if GetVarType(c)="Multi" then    
                                    f4.writeline("--"+c.name)
                                    f4.writeline(";with T"+Ctext(cnt)+" as (Select [Serial],"+x.name+","+a.name+","+b.name+",Codes from finaldata cross apply(values")                                        
                                    for each y1 in x.categories
                                        for each y2 in a.categories
                                            for each y3 in b.categories
                                                f4.writeline("('"+Lcase(y1.name)+"','"+Lcase(y2.name)+"','"+Lcase(y3.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"{"+y2.name+"}_"+b.name+"{"+y3.name+"}_"+c.name+"]),")
                                            next
                                        next    
                                    next
                                    f4.writeline(") x("+x.name+","+a.name+","+b.name+",Codes) where Len(Codes)>0")
                                    f4.writeline(")select Serial,"+x.name+","+a.name+","+b.name+",Ans as "+c.name+"Ans into Fact"+c.name+" from T"+Ctext(cnt)+" cross apply split(Codes,';')")
                                    f4.writeline("ALTER TABLE Fact"+c.name+" ALTER Column "+c.name+"Ans NVARCHAR(500)")
                                    f4.writeline("ALTER TABLE Fact"+c.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+c.name+" FOREIGN KEY ("+c.name+"Ans) REFERENCES Dim"+c.name+"("+c.name+"Ans)")
                                    
                                else
                                    f4.writeline("--"+c.name)
                                    f4.writeline("Select [Serial],"+x.name+","+a.name+","+b.name+","+c.name+" into Fact"+c.name+" from finaldata cross apply(values")                                        
                                    for each y1 in x.categories
                                        for each y2 in a.categories
                                            for each y3 in b.categories
                                                f4.writeline("('"+Lcase(y1.name)+"','"+Lcase(y2.name)+"','"+Lcase(y3.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"{"+y2.name+"}_"+b.name+"{"+y3.name+"}_"+c.name+"]),")
                                            next
                                        next    
                                    next
                                    f4.writeline(") x("+x.name+","+a.name+","+b.name+","+c.name+") where Len("+c.name+")>0")                                    
                                    if GetVarType(c)="Single" then f4.writeline("ALTER TABLE Fact"+c.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+c.name+" FOREIGN KEY ("+c.name+") REFERENCES Dim"+c.name+"("+c.name+")")
                                end if
                                f4.writeline("ALTER TABLE Fact"+c.name+" WITH NOCHECK ADD CONSTRAINT FK_Serial_"+c.name+" FOREIGN KEY (Serial) REFERENCES FinalData(Serial)")
                                
                                f4.writeline("ALTER TABLE Fact"+c.name+" ALTER Column "+b.name+" NVARCHAR(500)")
                                f4.writeline("ALTER TABLE Fact"+c.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+" FOREIGN KEY ("+b.name+") REFERENCES Dim"+b.name+"("+b.name+")")
                                f4.writeline("ALTER TABLE Fact"+c.name+" ALTER Column "+a.name+" NVARCHAR(500)")
                                f4.writeline("ALTER TABLE Fact"+c.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+"_"+b.name+" FOREIGN KEY ("+a.name+") REFERENCES Dim"+a.name+"("+a.name+")")
                                f4.writeline("ALTER TABLE Fact"+c.name+" ALTER Column "+x.name+" NVARCHAR(500)")
                                f4.writeline("ALTER TABLE Fact"+c.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+"_"+b.name+"_"+c.name+" FOREIGN KEY ("+x.name+") REFERENCES Dim"+x.name+"("+x.name+")")
                            end if
                        next
                    else 'level 2 b
                        f4.writeline(GetDim(b))
                        if GetVarType(b)="Multi" then                        
                            f4.writeline("--"+b.name)
                            f4.writeline(";with T"+Ctext(cnt)+" as (Select [Serial],"+x.name+","+a.name+",Codes from finaldata cross apply(values")                
                            for each y1 in x.categories
                                for each y2 in a.categories
                                    f4.writeline("('"+Lcase(y1.name)+"','"+Lcase(y2.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"{"+y2.name+"}_"+b.name+"]),")
                                next    
                            next                    
                            f4.writeline(") x("+x.name+","+a.name+",Codes) where Len(Codes)>0")                        
                            f4.writeline(")select Serial,"+x.name+","+a.name+",Ans as "+b.name+"Ans into Fact"+b.name+" from T"+Ctext(cnt)+" cross apply split(Codes,';')")
                            f4.writeline("ALTER TABLE Fact"+b.name+" ALTER Column "+b.name+"Ans NVARCHAR(500)")
                            f4.writeline("ALTER TABLE Fact"+b.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+b.name+" FOREIGN KEY ("+b.name+"Ans) REFERENCES Dim"+b.name+"("+b.name+"Ans)")
                        else
                            f4.writeline("--"+b.name)
                            f4.writeline("Select [Serial],"+x.name+","+a.name+","+b.name+" into Fact"+b.name+" from finaldata cross apply(values")                
                            for each y1 in x.categories
                                for each y2 in a.categories
                                    f4.writeline("('"+Lcase(y1.name)+"','"+Lcase(y2.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"{"+y2.name+"}_"+b.name+"]),")
                                next    
                            next                    
                            f4.writeline(") x("+x.name+","+a.name+","+b.name+") where Len("+b.name+")>0")                        
                            if GetVarType(b)="Single" then f4.writeline("ALTER TABLE Fact"+b.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+b.name+" FOREIGN KEY ("+b.name+") REFERENCES Dim"+b.name+"("+b.name+")")
                        end if
                        f4.writeline("ALTER TABLE Fact"+b.name+" WITH NOCHECK ADD CONSTRAINT FK_Serial_"+b.name+" FOREIGN KEY (Serial) REFERENCES FinalData(Serial)")
                        
                        f4.writeline("ALTER TABLE Fact"+b.name+" ALTER Column "+a.name+" NVARCHAR(500)")
                        f4.writeline("ALTER TABLE Fact"+b.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+" FOREIGN KEY ("+a.name+") REFERENCES Dim"+a.name+"("+a.name+")")
                        f4.writeline("ALTER TABLE Fact"+b.name+" ALTER Column "+x.name+" NVARCHAR(500)")
                        f4.writeline("ALTER TABLE Fact"+b.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+"_"+b.name+" FOREIGN KEY ("+x.name+") REFERENCES Dim"+x.name+"("+x.name+")")
                    end if
                next
            else  'level 1 a
                f4.writeline(GetDim(a))
                if GetVarType(a)="Multi" then 
                    f4.writeline("--"+a.name)
                    f4.writeline(";with T"+Ctext(cnt)+" as (Select [Serial],"+x.name+",Codes from finaldata cross apply(values")                
                    for each y1 in x.categories                
                        f4.writeline("('"+Lcase(y1.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"]),")
                    next
                    f4.writeline(") x("+x.name+",Codes) where Len(Codes)>0")                
                    f4.writeline(")select Serial,"+x.name+",Ans as "+a.name+"Ans into Fact"+a.name+" from T"+Ctext(cnt)+" cross apply split(Codes,';')")                    
                    f4.writeline("ALTER TABLE Fact"+a.name+" ALTER Column "+a.name+"Ans NVARCHAR(500)")
                    f4.writeline("ALTER TABLE Fact"+a.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+a.name+" FOREIGN KEY ("+a.name+"Ans) REFERENCES Dim"+a.name+"("+a.name+"Ans)")                
                else
                    f4.writeline("--"+a.name)
                    f4.writeline("Select [Serial],"+x.name+","+a.name+" into Fact"+a.name+" from finaldata cross apply(values")                
                    for each y1 in x.categories                
                        f4.writeline("('"+Lcase(y1.name)+"',["+x.name+"{"+y1.name+"}_"+a.name+"]),")
                    next
                    f4.writeline(") x("+x.name+","+a.name+") where Len("+a.name+")>0")          
                    if GetVarType(a)="Single" then f4.writeline("ALTER TABLE Fact"+a.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+a.name+" FOREIGN KEY ("+a.name+") REFERENCES Dim"+a.name+"("+a.name+")")                
                end if    
                f4.writeline("ALTER TABLE Fact"+a.name+" ALTER Column "+x.name+" NVARCHAR(500)")
                f4.writeline("ALTER TABLE Fact"+a.name+" WITH NOCHECK ADD CONSTRAINT FK_Serial_"+a.name+" FOREIGN KEY (Serial) REFERENCES FinalData(Serial)")                
                f4.writeline("ALTER TABLE Fact"+a.name+" WITH NOCHECK ADD CONSTRAINT FK_Dim_"+x.name+"_"+a.name+" FOREIGN KEY ("+x.name+") REFERENCES Dim"+x.name+"("+x.name+")")
            end if
        next
    else 'not loop
        if x.leveldepth=1 and GetVarType(x)="Multi" then
            f4.writeline("--"+x.name)
            f4.writeline("select Serial,Ans As "+x.name+"Ans into Fact"+x.name+" from [Finaldata] cross apply split("+x.name+",';')")
            f4.writeline("ALTER TABLE Fact"+x.name+" ALTER Column "+x.name+"Ans NVARCHAR(500)")
            f4.writeline("ALTER TABLE Fact"+x.name+" WITH NOCHECK ADD CONSTRAINT FK_Serial_"+x.name+" FOREIGN KEY (Serial) REFERENCES FinalData(Serial)")            
        end if
        if x.ObjectTypeValue=ObjectTypesConstants.mtVariable then
            f4.writeline(GetDim(x))
        end if    
    end if
end sub

function IsMulti(x)
    dim y
    y=false    
    if x.dataType=DataTypeConstants.mtCategorical then
         if x.maxvalue is null or x.maxvalue>1 then
             y=true
        end if
    end if
    IsMulti=y
end function

function GetDim(x)
    select case x.ObjectTypeValue 
        case ObjectTypesConstants.mtVariable 
            if x.dataType= DataTypeConstants.mtCategorical then
                    dim z1,z2
                    if GetVarType(x)="Multi" then
                        z2="Create Table Dim"+x.name+"("+x.name+"Id int IDENTITY(1,1) NOT NULL,"+x.name+"Ans nvarchar(500) NOT NULL PRIMARY KEY,"+x.name+"Code nvarchar(500),"+x.name+"Factor float,"+x.name+"Net1 nvarchar(100))"+mr.crlf
                        if x.leveldepth=1 then z2=z2+"ALTER TABLE Fact"+x.name+" WITH NOCHECK ADD CONSTRAINT FK_Fact_"+x.name+" FOREIGN KEY ("+x.name+"Ans) REFERENCES Dim"+x.name+"("+x.name+"Ans)"+mr.crlf
                    else
                        z2="Create Table Dim"+x.name+"("+x.name+"Id int IDENTITY(1,1) NOT NULL,"+x.name+" nvarchar(500) NOT NULL PRIMARY KEY,"+x.name+"Code nvarchar(500),"+x.name+"Factor float,"+x.name+"Net1 nvarchar(100))"+mr.crlf
                        if x.leveldepth=1 then z2=z2+"ALTER TABLE FinalData WITH NOCHECK ADD CONSTRAINT FK_FinalData_"+x.name+" FOREIGN KEY ("+x.name+") REFERENCES Dim"+x.name+"("+x.name+")"+mr.crlf
                    end if
                    for each z1 in x.categories
                        if GetVarType(x)="Multi" then
                            z2=z2+"insert into Dim"+x.name+"("+x.name+"Ans,"+x.name+"Code) values('"+Lcase(z1.name) +"','"+ replace(z1.label,"'","")+"');"+mr.crlf                            
                        else
                            z2=z2+"insert into Dim"+x.name+"("+x.name+","+x.name+"Code) values('"+Lcase(z1.name) +"','"+ replace(z1.label,"'","")+"');"+mr.crlf                            
                        end if
                    next
                    
                    GetDim=z2                
            end if            
        case objectTypesconstants.mtarray
            dim z3,z4
            z4="Create Table Dim"+x.name+"("+x.name+"Id int IDENTITY(1,1) NOT NULL,"+x.name+" nvarchar(500) NOT NULL PRIMARY KEY,"+x.name+"Code nvarchar(500),"+x.name+"Factor float,"+x.name+"Net1 nvarchar(100))"+mr.crlf
            for each z3 in x.categories
                z4=z4+"insert into Dim"+x.name+"("+x.name+","+x.name+"Code) values('"+Lcase(z3.name) +"','"+ replace(z3.label,"'","")+"');"+mr.crlf                            
            next
            GetDim=z4            
        end select
end function

function GetName(x)
    GetName=x.Replace(".","_").replace("[","").replace("]","")
end function
