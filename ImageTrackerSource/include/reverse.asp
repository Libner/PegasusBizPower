<% DIM numhebstr_br ' number of output strings after Hebrew()

Function reverse(word) 
dim revword,engword,i,letter,num,eng,heb,pos,newword
  if  word<>nil AND WORD<>"" then
	revword=""
	engword=""
    heb=" ,./':-()%"+chr(34)
    eng=".,?!'()"+chr(34)
    postcheck="."
    pos=0
    i=1
	do while i<= len(word)
	   letter=Mid(word, i, 1)
       num=asc(letter)
       if (num>=224 and num<=250) then 
          exit do
       end if
       i=i+1
    loop
    if i<=len(word) then 
        for i=1 to len(word)
		    letter=Mid(word, i, 1)
            num=asc(letter)
		    if (num>=224 and num<=250) or (instr(1,heb,letter)<>0 and pos=0) then
'hebrew
                  if letter="(" then  
                      letter=")" 
                  elseif letter=")" then  
                      letter="(" 
                  end if
                  if pos>0 then revword=" " & revword 
                      revword=letter & revword
                      pos=0
                  else
'english
                  if instr(1,eng,letter)=0 or (instr(1,postcheck,letter)<>0 and mid(word,i+1,1)<>" " and mid(word,i+1,1)<>"") then
                      revword=left(revword,pos) & letter & mid(revword,pos+1)
                      pos=pos+1
                  else
'hebrew
                      if letter="(" then  
                         letter=")" 
                      elseif letter=")" then  
                         letter="(" 
                      end if
                      revword=letter & revword
                      pos=0
                  end if
           end if
      next	
      reverse=revword
   else 
      reverse=word
   end if
 else
    reverse=""
 end if
 end Function

 Function hebrew(word,width) 
	 dim place,newstring,newword,pos,pos1,insstr,spflag,entpos
	
	 numhebstr_br=0
	 if word<>nil AND WORD<>"" then
	    width=width+1
        word=trim(word)
        spflag=false
        for pos=1 to len(word)
           if mid(word,pos,1)<>" " then
              newword=newword & mid(word,pos,1)
              spflag=false
           else
              if spflag=false then
                  newword=newword & " "
                  spflag=true
              end if   
           end if
        next
        word=newword
        newstring=""
	    do while len(word)>0
		   numhebstr_br=numhebstr_br+1
           if width<len(word) then
    		   place=width
		       do while mid(word,place,1)<>" " and place>1 
			      place=place-1
		       loop
		       if place=1 then 
                   place=width
                   do while place<=len(word) 
                      if mid(word,place,1)=" " then exit do
                         place=place+1
                      loop
                   end if
               else
                  place=len(word)+1
               end if
               entpos=instr(word,chr(13)&chr(10))
               if entpos<>0 and entpos<=place then
                  word=left(word,entpos-1)&" "&mid(word,entpos+1)
                  place=entpos
               end if
               newword=reverse(left(word,place-1))
               pos=1
               do while pos<>0
                  pos=InStr(pos,newword,"http://",1)
                  if pos<>0 then
                     pos1=InStr(pos,newword," ",1)-1
                     if pos1 < 0 then pos1=len(newword)
                        v_str=mid(newword,pos,pos1-pos+1)
                        v_pos=InStr(8,v_str,"/",1)			
		                if v_pos<>0 then 
                            v_str0=left(v_str,v_pos-1)+"..."
                        else
                            v_str0=v_str
                        end if 
                        insstr="<a href='"+v_str+"'>"+v_str0+"</a>"
                        newword=left(newword,pos-1)+insstr+mid(newword,pos1+1)
                        pos=pos+len(insstr)
                     end if
               loop
               pos=1
               do while pos<>0
                  pos=InStr(pos,newword,"www.",1)
                  if pos<>0 then
                     spFlag=true
                     if pos>=7 then
                        if mid(newword,pos-7,7)="http://" then
                           pos=pos+1
                           spFlag=false
                        end if
                     end if
                     if spFlag=true then
                        pos1=InStr(pos,newword," ",1)-1
                        if pos1 < 0 then pos1=len(newword)
                        v_str=mid(newword,pos,pos1-pos+1)
		                v_pos=InStr(1,v_str,"/",1)			
		                if v_pos<>0 then 
                            v_str0=left(v_str,v_pos-1)+"..."
                        else
                            v_str0=v_str
                        end if 
 				        insstr="<a href='http://"+v_str+"'>"+v_str0+"</a>"
                        newword=left(newword,pos-1)+insstr+mid(newword,pos1+1)
                        pos=pos+len(insstr)
                     end if 
                  end if
               loop
               pos=1
               do while pos<>0
                  pos=InStr(pos,newword,"@",1)
                  if pos<>0 then
                      do while pos>0
					     if mid(newword,pos,1)=" " then
					         exit do
					     end if
                         pos=pos-1
                      loop
                      pos=pos+1
                      pos1=InStr(pos,newword," ",1)-1
                      if pos1<=0 then pos1=len(newword)
                      insstr="<a href='mailto:"+mid(newword,pos,pos1-pos+1)+"'>"+mid(newword,pos,pos1-pos+1)+"</a>"
                      newword=left(newword,pos-1)+insstr+mid(newword,pos1+1)
                      pos=pos+len(insstr)
                  end if
               loop  
               newstring=newstring & newword 
		       word=right(word,len(word)-place+1)
               word=trim(word) 
               if len(word)>0 then newstring=newstring & "<br>" 
           loop
           newstring=newstring & reverse(word)
           hebrew=newstring
       else
           hebrew=""
       end if
end Function

Function hebrewDate(date)
  if date=nil then
      hebrewDate=""
  else
      hebrewDate=CStr(Day(date))&"/"&CStr(Month(date))&"/"&CStr(Year(date))
  end if
end Function

'SQL Server
Function sFix(word)
   if word<>nil AND WORD<>"" then
       sFix=Replace(word,"'","'+char(39)+'")
   else
       sFix=""
   end if
end Function
   
Function vFix(word)
  if word<>nil AND WORD<>"" then
     word=Replace(word,"<","&lt;")
     word=Replace(word,">","&gt;")
     word=Replace(word,chr(34),"&quot;")
     vFix=word
  else
     vFix=""
  end if
end Function

   Function Encode(sIn)
            Dim x, y, abfrom, abto
            Encode = "" : abfrom = ""

            For x = 0 To 25 : abfrom = abfrom & Chr(65 + x) : Next
            For x = 0 To 25 : abfrom = abfrom & Chr(97 + x) : Next
            For x = 0 To 9 : abfrom = abfrom & CStr(x) : Next

            abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
            For x = 1 To Len(sIn) : y = InStr(abfrom, Mid(sIn, x, 1))
                If y = 0 Then
                    Encode = Encode & Mid(sIn, x, 1)
                Else
                    Encode = Encode & Mid(abto, y, 1)
                End If
            Next
    End Function
    
    Function Decode(sIn)
            Dim x, y, abfrom, abto
            Decode = "" : abfrom = ""

            For x = 0 To 25 : abfrom = abfrom & Chr(65 + x) : Next
            For x = 0 To 25 : abfrom = abfrom & Chr(97 + x) : Next
            For x = 0 To 9 : abfrom = abfrom & CStr(x) : Next

            abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
            For x = 1 To Len(sIn) : y = InStr(abto, Mid(sIn, x, 1))
                If y = 0 Then
                    Decode = Decode & Mid(sIn, x, 1)
                Else
                    Decode = Decode & Mid(abfrom, y, 1)
                End If
            Next
     End Function
%>