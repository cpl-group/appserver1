<%@Language="VBScript"%>
<% 
    folder = "\\10.0.7.2\genergy" 
 
    set fso = server.createobject("Scripting.fileSystemObject") 
    set fold = fso.getFolder(folder) 
    fileCount = fold.files.count 
    dim fNames() 
    redim fNames(fileCount) 
    cFcount = 0 
    for each file in fold.files 
        cFcount = cFcount + 1 
        fNames(cFcount) = lcase(file.name) 
    next 
    for tName = 1 to fileCount 
        for nName = (tName + 1) to fileCount 
            if strComp(fNames(tName),fNames(nName),0)=1 then 
                buffer = fNames(nName) 
                fNames(nName) = fNames(tName) 
                fNames(tName) = buffer 
            end if 
        next 
    next 
    for i = 1 to fileCount 
        content = content & fNames(i) & "<br>" 
    next 
    Response.Write content 
%>