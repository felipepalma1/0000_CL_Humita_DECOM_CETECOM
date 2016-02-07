Option Explicit

Dim a

a=0

Do Until a=5
a=a+1
MsgBox a
Loop


Dim b

b=0

Do Until b>5
b=b+1
MsgBox b
Loop

Dim c

c=0

Do Until c<5
c=c+1
MsgBox c
Loop


Dim d

d=0

Do Until d=5 Or d<5
d=d+1
MsgBox d
Loop

Dim e

e = 0

Do
 e=e+1
 MsgBox e
Loop until e=5
