%===============Data Input===============%
global Grid Branch Node PMSG w0

%% //////////UPLOAD DATA FROM EXCEL//////////%
[filename,pathname] = uigetfile('*.xlsx','Please select data');
Filename=strcat(pathname,filename);
Grid=xlsread(Filename,'Network','A2:E2');
Branch=xlsread(Filename,'Branch',strcat('A2:F',int2str(Grid(1,2)+1)));     % upload branch data
Node=xlsread(Filename,'Node',strcat('A2:M',int2str(Grid(1,1)+1)));      % upload node data
PMSG=xlsread(Filename,'PMSG',strcat('A2:Z',int2str(Grid(1,5)+1)));
N_PMSG=Grid(5);% number of PMSGs in the wind farm;
Node_num=Grid(1);% number of the network node£»
w0=2*3.14*50*ones(N_PMSG,1); % synchrounous angle frequency in vector form
%% PI parameters of the PMSGs
% DC voltage outer control loop
Kp_Vdc=PMSG(:,17);Ki_Vdc=PMSG(:,18);
% Reactive power outer control loop
Kp_Q=PMSG(:,19);Ki_Q=PMSG(:,20);
% Active current inner control loop
Kp_Id=PMSG(:,21);Ki_Id=PMSG(:,22);
% Reactive current inner control loop
Kp_Iq=PMSG(:,23);Ki_Iq=PMSG(:,24);
% PLL
Kp_pll=w0.*PMSG(:,25);Ki_pll=w0.*PMSG(:,26);
%% Filter capacitance and inductance in the dc-link
Cdc=PMSG(:,16)./w0;
Xf=PMSG(:,15);
%% Reference value of control loop
Vdc_ref=PMSG(:,14);% dc voltage reference
% Reference of reactive power control loop
Qo_ref=zeros(N_PMSG,1);
% Constant active power source of the MSC
Po_ref=zeros(N_PMSG,1);
% Id_fault and Iq_fault current injections commands during the fault
Id_fault=zeros(N_PMSG,1);Iq_fault=Id_fault;
for k=1:Node_num
    if Node(k,11)==3 % identify the type of PMSG
        Po_ref(Node(k,10))=Node(k,4);% PMSG's active power
        Qo_ref(Node(k,10))=Node(k,5);% PMSG's reactive power
        Id_fault(Node(k,10))=Node(k,12);% Id_fault: active current injection during the fault
        Iq_fault(Node(k,10))=Node(k,13);% Iq_fault: reactive current injection during the fault
    end
end
%% windfarm's network matrix : Mnet£»PMSG's head node number:#1£¬#2,...#N£»PCC node number:#N+1;Idea bus node number: #N+2;
x=zeros(N_PMSG,N_PMSG);
XL=Branch(Grid(2),5);% reactance of AC transmission line
RL=Branch(Grid(2),4);% resistance of AC transmission line
Mnet_XL=XL*ones(N_PMSG,N_PMSG);
Mnet_RL=RL*ones(N_PMSG,N_PMSG);
for k=1:N_PMSG
    x(k,k)=Branch(k,5);
end
Mnet_XL=Mnet_XL+x;



