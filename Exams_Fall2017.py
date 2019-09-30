#
def finalExam(filename):
    ans = []
    qLabel = []
    point = []
    
    # how to check question 1 or 2 :D
    ans.append("full-mark")
    qLabel.append("1")
    point.append(1)
    
    ans.append("full-mark")
    qLabel.append("2")
    point.append(1)

    ans.append(grade("theme"))
    qLabel.append("3")
    point.append(1)

    ans.append(grade("sheet1","bold&size","a",1,"k",1))
    qLabel.append("4.a")
    point.append(2)

    ans.append(grade("sheet1","colW",9)) # other columns are not created yet
    qLabel.append("4.b")
    point.append(1)

    ans.append(grade("sheet1","colHide"))
    qLabel.append("4.c")
    point.append(2)

    ans.append(grade("sheet1","sheetName"))
    qLabel.append("4.d")
    point.append(1)

    ans.append(grade("sheet1","sheetProtect"))
    qLabel.append("4.e")
    point.append(1)

    ans.append(grade("sheet2","sheetName"))
    qLabel.append("5.a")
    point.append(1)

    ans.append(grade("sheet2","freeze"))
    qLabel.append("5.b")
    point.append(1)

    ans.append(grade("sheet2","conFormat","H")) # if the grades are the same it can be written in one command
    qLabel.append("6")
    point.append(1)

    ans.append(grade("sheet2","conFormat","G")) 
    qLabel.append("7")
    point.append(1)

    ans.append(grade("sheet2","value","a",6261,"a",6264)) 
    qLabel.append("8")
    point.append(1)
    
    ans.append(grade("sheet2","bold&size","a",6261,"a",6264)) 
    qLabel.append("8.a")
    point.append(2)

    ans.append(grade("sheet2","formulaE","b",6261,"f",6264)) 
    qLabel.append("8.b")
    point.append(4)

    ans.append(grade("sheet2","border","a",6261,"f",6264)) 
    qLabel.append("8.c")
    point.append(1)

    ans.append(grade("sheet2","format","c",6262,"f",6262)) 
    qLabel.append("8.d")
    point.append(1)

    ans.append(grade("sheet2","value","j",1)) 
    qLabel.append("9.a")
    point.append(1)

    ans.append(grade("sheet2","bold&size","j",1)) 
    qLabel.append("9.b")
    point.append(1)

    ans.append(grade("sheet2","formulaF","j",2,"j",6260)) 
    qLabel.append("9.c")
    point.append(1)

    ans.append(grade("sheet2","format","j",2,"j",6260))
    qLabel.append("9.d")
    point.append(1)

    ans.append(grade("sheet2","value","k",1)) 
    qLabel.append("10.a")
    point.append(1)

    ans.append(grade("sheet2","formulaFValue","k",2,"k",6260,"expensive")) 
    qLabel.append("10.b")
    point.append(2)
    
    ans.append(grade("sheet2","formulaFValue","k",2,"k",6260,"inexpensive")) 
    qLabel.append("10.c")
    point.append(2)
    
    ans.append(grade("sheet2","formulaFValue","k",2,"k",6260,"affordable")) 
    qLabel.append("10.d")
    point.append(2)

    ans.append(grade("sheet2","value","l",1)) 
    qLabel.append("11.a")
    point.append(1)

    ans.append(grade("sheet2","formulaF","l",2,"l",6260)) 
    qLabel.append("11.b")
    point.append(1)

    ans.append(grade("sheet2","format","l",2,"l",6260)) 
    qLabel.append("11.c")
    point.append(1)

    ans.append(grade("sheet3","sheetName"))
    qLabel.append("12.a")
    point.append(1)

    ans.append(grade("sheet3","formulaValue","l",2,"l",6260))
    qLabel.append("12.b")
    point.append(2)

    ans.append(grade("sheet3","conFormat","D")) 
    qLabel.append("12.c")
    point.append(1)

    ans.append(grade("sheet3","value","n",1))
    qLabel.append("13.a")
    point.append(1)

    ans.append(grade("sheet3","bold&size","n",1))
    qLabel.append("13.b")
    point.append(1)

    ans.append(grade("sheet3","formulaF","n",2,"n",6260)) 
    qLabel.append("13.c")
    point.append(2)
    
    resTable = printResult(qLabel,point,ans,filename)
    return resTable

