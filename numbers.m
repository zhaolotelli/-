%import data
d = 18;
P = xlsread('numbers.xlsx', 'B1:D18');
Q = xlsread('numbers.xlsx', 'E1:N18');

%compute max number
maxP = max(P, [], 'all');
maxQ = max(Q, [], 'all');
m = max(maxP, maxQ);

%generate input
f = 999*ones(1, d*m);
[rowP, colP] = size(P);
for i = 1:rowP
    for j = 1:colP
        if ~isnan(P(i,j))
            f((i-1)*m+P(i,j)) = -1;
        end
    end
end
[rowQ, colQ] = size(Q);
for i = 1:rowP
    for j = 1:colP
        if ~isnan(Q(i,j))
            f((i-1)*m+Q(i,j)) = 0;
        end
    end
end

intcon = 1:(d*m);

A = [];
for i = 1:d
    A = [A, eye(m)];
end

b = ones(1, m);

Aeq = [];
for i = 1:d
    Aeq = [Aeq, [zeros(i-1,m);ones(1,m);zeros(d-i,m)]];
end

beq = ones(1, d);

lb = zeros(1, d*m);

ub = ones(1, d*m);

%evaluate
x = intlinprog(f,intcon,A,b,Aeq,beq,lb,ub);

%output
x_reshaped = reshape(x, [m, d]);
[~,number] = max(x_reshaped, [], 1);
xlswrite('numbers.xlsx',number','O1:O18');