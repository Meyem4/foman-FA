%   This file is part of DSSLab. Copyright (C) 2020-2021 Meryem Tahri and Haytham Tahri




function varargout = projet(varargin)
% PROJET MATLAB code for projet.fig
%      PROJET, by itself, creates a new PROJET or raises the existing
%      singleton*.
%
%      H = PROJET returns the handle to a new PROJET or the handle to
%      the existing singleton*.
%
%      PROJET('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PROJET.M with the given input arguments.
%
%      PROJET('Property','Value',...) creates a new PROJET or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before projet_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to projet_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help projet

% Last Modified by GUIDE v2.5 13-Oct-2020 15:26:38



% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @projet_OpeningFcn, ...
                   'gui_OutputFcn',  @projet_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before projet is made visible.
function projet_OpeningFcn(hObject, ~, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to projet (see VARARGIN)

% Choose default command line output for projet
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes projet wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = projet_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function listbox1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
n = get(handles.popupmenu1,'Value');
files = getappdata(0,'pushbutton1');
sz = size(files,2);
for i =1:sz
    eval(['A' num2str(i) ' = xlsread(char(files(i)))']);
    v = size(eval(['A' num2str(i)]),2);
    if n > v-1
        f = msgbox(['Invalid Value for ' char(files(i))], 'Error','error');
    else
        setappdata(0,'popupmenu1',n);
    end
end



% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
[filename dirname] = uiputfile; %set file name and location
fullname = fullfile(dirname,filename); %make full file name
saveas(gcf,fullname); %save plot as fullname


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
[filename pathname] = uigetfile('*.xlsx', 'Choose files to load:','MultiSelect','on');
guidata(hObject, handles);
set(handles.listbox1, 'string', filename);
files = strcat(pathname,filename);
setappdata(0,'pushbutton1',files);
nfile = numel(string(filename));
if nfile < 2
    msgbox('Minimum 2 matrix', 'Error','error');
end


% --------------------------------------------------------------------
function Help_Callback(hObject, eventdata, handles)
% hObject    handle to Help (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
open Help.pdf



function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double
alpha1 = str2double(get(hObject,'string'));
if isnan(alpha1) 
alpha = 0.5;
else
alpha = str2double(get(handles.edit1,'String'));
end
mu = 1 - alpha; 


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

files = getappdata(0,'pushbutton1');
sz = size(files,2);
v = getappdata(0,'popupmenu1');
n = v+1;
alpha1 = str2num(get(handles.edit1,'String'));
alpha1
if isempty(alpha1) 
alpha = 0.5;
else
alpha = alpha1;
alpha
end
mu = 1 - alpha;

for i =1:sz
    eval(['A' num2str(i) ' = xlsread(char(files(i)))']);    
end
  
l = A1;
m_temp = A1;
u = A1;

for k = 1:sz-1       
            B = eval(['A' num2str(k+1)]);            
    for i = 1:n
        for j = 1:n    
            l(i,j) = min([l(i,j),B(i,j)]);
            m_temp(i,j) = (m_temp(i,j)*B(i,j));            
            u(i,j) = max([u(i,j),B(i,j)]);
        end
    end
end

for i = 1:n
        for j = 1:n
            m(i,j) = (m_temp(i,j)^(1/sz));
        end
end

switch v+1
    case 3
    cr = 0.58
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3)];
    end
    case 4
	cr = 0.90
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3),l(i,4),m(i,4),u(i,4)];
    end
    case 5
	cr = 1.12
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3),l(i,4),m(i,4),u(i,4),l(i,5),m(i,5),u(i,5)];
    end
    case 6
    cr = 1.24
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3),l(i,4),m(i,4),u(i,4),l(i,5),m(i,5),u(i,5),l(i,6),m(i,6),u(i,6)];
    end
    case 7
	cr = 1.32
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3),l(i,4),m(i,4),u(i,4),l(i,5),m(i,5),u(i,5),l(i,6),m(i,6),u(i,6),l(i,7),m(i,7),u(i,7)];
    end
    case 8
	cr = 1.41
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3),l(i,4),m(i,4),u(i,4),l(i,5),m(i,5),u(i,5),l(i,6),m(i,6),u(i,6),l(i,7),m(i,7),u(i,7),l(i,8),m(i,8),u(i,8)];
    end
    case 9
    cr = 1.45
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3),l(i,4),m(i,4),u(i,4),l(i,5),m(i,5),u(i,5),l(i,6),m(i,6),u(i,6),l(i,7),m(i,7),u(i,7),l(i,8),m(i,8),u(i,8),l(i,9),m(i,9),u(i,9)];
    end
    case 10
	cr = 1.49
    for i=1:n
        A(i,:)=[l(i,1),m(i,1),u(i,1),l(i,2),m(i,2),u(i,2),l(i,3),m(i,3),u(i,3),l(i,4),m(i,4),u(i,4),l(i,5),m(i,5),u(i,5),l(i,6),m(i,6),u(i,6),l(i,7),m(i,7),u(i,7),l(i,8),m(i,8),u(i,8),l(i,9),m(i,9),u(i,9),l(i,10),m(i,10),u(i,10)];
    end
end

alpha
mu
for i = 1:n
    for j = 1:n
        if i <= j
        Ag(i,j) = mu*((m(i,j)-l(i,j))*alpha+l(i,j))+(1-mu)*(u(i,j)-(u(i,j)-m(i,j))*alpha); 
        else
        Ag(i,j) = 1/(mu*((m(j,i)-l(j,i))*alpha+l(j,i))+(1-mu)*(u(j,i)-(u(j,i)-m(j,i))*alpha));
        end
    end
end
Vpmax = max(eig(Ag));
A
Ag
set(handles.uitable1,'Data',Ag);
Vpmax
CI = (Vpmax-n)/(n-1);
CR = CI/cr
if CR < 0.1
    txt = 'Data are consistent';
    set(handles.text6,'String',txt);
    set(handles.text6,'ForegroundColor','[0 0.5 0]');
else
    txt = 'Revise the judgment matrix';
    set(handles.text6,'String',txt);
    set(handles.text6,'ForegroundColor','red');
end
set(handles.text3,'String',CR)
[V,D] = eig(Ag) %produces a diagonal matrix D of eigenvalues and a full matrix V whose columns are the corresponding eigenvectors
W = V(:,1)% In this case
Wn = W./sum(W)

%%%% Weight ranking %%%%
[R,TIEADJ] = tiedrank(-Wn)
T = [Wn R]
set(handles.uitable2,'Data',T);


%%% Plot bar %%%%

x = 1:n; % arbitrary array
y = Wn*[100];

% Create a vector of '%' signs

   pct = char(ones(size(y,1),1)*'%');

% Append the '%' signs after the percentage values

   new_yticks = [char(y),pct];

% 'Reflect the changes on the plot

   set(gca,'yticklabel',new_yticks)

bar(x,y);

ylim([0 100])

xlabel('Criteria');

ylabel('Percentage');

labels = arrayfun(@(value) num2str(value,'%2.2f'),y,'UniformOutput',false);
z = '%';
txt = strcat(labels,z);%concatener y avec %
text(x,y,txt,'HorizontalAlignment','center','VerticalAlignment','bottom')

%%%% Export bar chart in pdf file %%%%
ax = gca;
exportgraphics(ax,'BarChart.pdf','ContentType','vector')

box off
filename = 'C:\Users\admin\Desktop\MATLAB\Output\resultat_A.xlsx';
writematrix(A,filename,'Sheet',1,'Range','A2:AC13')
writematrix(Ag,filename,'Sheet',1,'Range','A15:AC27')
