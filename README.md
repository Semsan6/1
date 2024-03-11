[nums,textd,rawd] = xlsread("BOOK.xlsx",'D2:D169'); %Открытие xlsx фала для чтения
p0=nums(~isnan(nums));
a=size(p0,1)-24; %Определяет размер массива
p=zeros(24,a);%Заполняет массив нулями
t=zeros(1,a);%Заполняет массив нулями
for i=1:a %Разбиение данных на группы из четырех элементов
    for j=1:24
    p(j,i)=[p0(i+j-1)]';
    end
t(i)=p0(i+24);
end
%load("my_neural_net1.mat");%Загрузка прогресса уже обученной ранее НС
net=newff(minmax(p),[60,24,12,1],{'logsig','logsig','logsig','purelin'}, 'trainlm'), %Активационная функция где p-множество входных векторов,а-количесвто входов НС,1-количество выходов НС,logsig-передаточная функция,purelin-вычисляет выход слоя от сетевого входа,trainlm-выполняет обучение мно-гослойной НС методом Ле-венберга-Марквардта
net.trainparam.show=10, %Количество эпох между показами
net.trainparam.epochs=1000, %Максимальное количество эпох обучения
net.trainparam.goal=1, %Целевое значение ошибки
net.trainparam.min_grad=1e-2, %Минимальное значение градиента
%net.trainparam.mu_max=1e+30;
[net,tr]=train(net,p,t); %Запуск процесса обучения НС

result_test=sim(net,p); %Прогноз на основе начальных данных
t1=((result_test-t)./t)*100;
result1=[t' result_test' t1']; %Вывод изначальных данных, рас-считанные данные с помощью НС,абсолютная погрешность
%result2=[tt' result_te' (((result_te-tt)./tt)*100)']; %Вывод рассчитанных данных с помощью НС, спрогнозированные данные на основе рассчитанных ранее дан-ных,абсолютная погрешность
for i=1:24
    prognoz(i)=p0(a+i);
end
pro=zeros(1,24);
for i=25:48 %Расчет данных на 24 часа вперед
    for j=1:24
    pro(j)=prognoz(i-25+j);
    end
prognoz(i)=sim(net,pro'); 
end 
prognozday=[prognoz']; %Вывод следующих 24 часов нагрузки
for i=1:a+4
haorses(i)=datetime(2023,10,1,i-1,0,0);
end
date11=datestr(haorses,'HH:MM');
for i=1:28
hours(i)=datetime(2023,10,2,i-5,0,0);
end
date21=datestr(hours,'HH:MM');
data1=cellstr(date21);
data=cellstr(date11);
col_head0={'Время'};%Вывод всех результатов в Excel
col_head1={'Начальные данные','Прогноз','Ошибка,%'};
col_head2={'Прогноз','Прогноз по прогнозу','Ошибка,%'};
col_head3={'Время','Прогнозируемое потребление на следующие 24 часа'};
xlswrite('data1.xlsx',col_head0,'Sheetl','A1');
xlswrite('data1.xlsx',data,'Sheetl','A2');
xlswrite('data1.xlsx',col_head1,'Sheetl','B1');
xlswrite('data1.xlsx',result1,'Sheetl','B6');
%xlswrite('data1.xlsx',col_head2,'Sheetl','F1');
%xlswrite('data1.xlsx',result2,'Sheetl','F10');
xlswrite('data1.xlsx',col_head3,'Sheetl','J1');
xlswrite('data1.xlsx',data1,'Sheetl','J2');
xlswrite('data1.xlsx',prognozday,'Sheetl','K2');

save('my_neural_net.mat','net');%Сохранение прогресса обучения НС
