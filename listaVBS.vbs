dim vetor (25)
i = 0
for each i in (vetor)
	vetor (i) = inputbox ("VALOR " & i)
next

pares = 0
impares = 0
soma = 0
maior = 0
menor = 0
negativos = 0

for each i in (vetor)
	if i mod 2 = 0 then
		pares = pares + 1
	else
		impares = impares + 1
	end if

	soma = soma + i

	if (i > maior) then
		maior = i
	end if 

	if (i < menor) then
		menor = i
	end if

	if i < 0 then
		negativos = 1
	end if
next 

media = soma / (vetor(i) + 25)

set FSO = createObject ("Scripting.FileSystemObject")

 const arquivo = "texto.csv"

 set arq = FSO.createTextFile(arquivo, 2)
 ' 1 leitura , 2 escrita/sobrescrita, 8 append

    arq.writeline "MEDIA ;" & media 
	arq.writeline "NUMEROS PARES ;" & pares 
	arq.writeline "NUMEROS IMPARES ;" & impares 
	arq.writeline "MEDIA ARITMETICA ;" & media 
	arq.writeline "MAIOR VALOR ;" & maior 
	arq.writeline "MENOR VALOR ;" & menor 
	arq.writeline "SOMATORIO ;" & soma 
	arq.writeline "NEGATIVOS ;" & negativos
	
	wscript.echo "arquivo gerado"

 arq.close
