clc
clear 
close all
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
filename = 'Book1.xlsx';
sheet = 1;
xlRange = 'A4:C27651';
St = xlsread(filename,sheet,xlRange);
[n,m]=size(St);
s=zeros(4456,6663);
for i=1:n
    s(St(i,1),St(i,2))=St(i,3);
end
[m,n]=size(s);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Lt=xlsread(filename,'A43736:C59813');
[nl,ml]=size(Lt);
l1=zeros(6663,30);
for i=1:nl
    l1(Lt(i,1),Lt(i,2))=Lt(i,3);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Ut=xlsread(filename,'A27655:C43732');
[nu,mu]=size(Ut);
u1=zeros(6663,30);
for i=1:nu
    u1(Ut(i,1),Ut(i,2))=Ut(i,3);
end
[n,d]=size(u1);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
v=zeros(6663,30);
for i=1:d
u=u1(1:6663,i);
l=l1(1:6663,i);
Aeq=[];
beq=[];
b=zeros(4456,1);
f=zeros(6663,1);
x=linprog(f,Aeq,beq,s,b,l,u);
% x=quadprog([],f,Aeq,beq,s,b,l,u);
v(:,i)=x;
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% k=0;
% for i=1:n
%     if norm(v(i,:))<10^-16
%         k=k+1;
%     end
% end
% k
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
k=0;
for i=1:n
    if sum(v(i,:)==0)==d
        k=k+1;
    end
end
    k
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
filename = 'output.xlsx';
A = v;
xlswrite(filename,A)
