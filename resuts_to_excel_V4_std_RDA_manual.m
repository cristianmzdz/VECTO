%% Generar Spreadsheet excel desde vmod
% Leer los datos del archivo CSV como una tablazczxczxczxc
clc
clear
global Summary_filename sheetname     ; %#ok<GVMIS>
% Solicita la ruta de la carpeta al donde esta los resultados de VECTO
try
    %% Leer el archivo vsum y exportarlo a una hoja de excel
    [Summary_file, folderPath]= uigetfile('*.vsum','Seleccionar archivo *.vsum','MultiSelect', 'on');
    % Verifica si se seleccionó una carpeta o si el usuario canceló
    if folderPath == 0
        disp('No se seleccionó ninguna carpeta. Se usara la carpeta actual por defecto.');
        return; % Sale del script si no se selecciona una carpeta
    else
        disp(['Carpeta seleccionada: ', folderPath]);
        root = cd (folderPath);
    end
    Summary_filename=[erase(Summary_file,'.vsum'),'_summary.xlsx'];
    readfile = readlines(Summary_file);
    Vecto_version = readfile{1};
    Hash = readfile{end-1};
    % Escribir los datos en un nuevo archivo CSV con ',' como delimitador
    writematrix(Vecto_version,Summary_filename,'Sheet','Summary','Range','A1',"AutoFitWidth",true,"PreserveFormat",true);
    writematrix(Hash,Summary_filename,'Sheet','Summary','Range','A2',"AutoFitWidth",true,"PreserveFormat",true);
    data_table = readtable(Summary_file,'VariableNamingRule', 'preserve','FileType','delimitedtext','delimiter',',');
    % Summary_data = table2cell(rows2vars(data_table(1:end-1,:)));
    Summary_data = table2array(rows2vars(data_table(1:end-1,:)));
    writecell(Summary_data,Summary_filename,'Sheet','Summary','Range','A3',"AutoFitWidth",true,"PreserveFormat",true);
    %% Leer el archivo vmod e importar las variables
    files={}; 
    files = uigetfile('*.vmod','Seleccionar archivo *.vmod','MultiSelect', 'on');
    % files = {dir('*.vmod').name};
    columnSelection = []; % Variable para almacenar la selección de columnas
    for i=1:length(files)
        file = files{i};
        sheetname = matlab.lang.makeValidName(file(1:min(26,end)));
        data = readtable(file,'VariableNamingRule', 'preserve','FileType','delimitedtext');
        timetable_data = table2timetable(data,"RowTimes",seconds(data.("time [s]")));
        timetable_data = removevars(timetable_data,'time [s]');
        if i == 1    % Si es la primera iteración
            % Obtener los nombres de las columnas
            columnNames = data.Properties.VariableNames;
            % Crear una lista de selección con los nombres de las columnas
            [indx,tf] = listdlg('PromptString','Seleccionar variables a exportar','ListString',columnNames,'SelectionMode','multiple','ListSize',[350 450]);
            % Si el usuario hace una selección, actualizar los datos para incluir solo las columnas seleccionadas
            if tf == 1
                columnSelection = indx;
            end
            %Preguntar si generar histogramas
            options = {'SI', 'NO'};
            [gen_histogram, ~] = listdlg('PromptString','¿Deseea generar histogramas?','ListString',options,'SelectionMode', ...
                'single','ListSize',[360 50]);
        end
        % Aplicar la selección de columnas y guardarlas en una tabla por cada archivo
        Files_data{i}= data(:,columnSelection); %#ok<*SAGROW>
        % Escribir los datos en un nuevo archivo CSV con ',' como delimitador
        writetable( Files_data{i}, Summary_filename,"Sheet",sheetname,"AutoFitWidth",true,"PreserveFormat",true);
        % Generar histogramas
        if gen_histogram == 1
            % Obtener los nombres de las columnas
            histogram_columnNames = Files_data{i}.Properties.VariableNames;
            if i == 1     % Crear una ventana emergente para introducir los valores en la primera iteracion
                %Preguntar si exportar los histogramas
                [excel_hist, ~] = listdlg('PromptString','¿Deseea exportar los histogramas a excel?','ListString',options,'SelectionMode', ...
                    'single','ListSize',[550 50]);
            end
            if i == 1     % Crear una ventana emergente para introducir los valores en la primera iteracion
                % Crear una lista de selección con los nombres de las columnas para X
                [indx_X,tf_X] = listdlg('PromptString',{'Seleccione la variable para el eje X (horizontal inferior) [→]','i.e : Speed'},'ListString',histogram_columnNames,'SelectionMode','single','ListSize',[350 550]);
                % Crear una lista de selección con los nombres de las columnas para Y, excluyendo la variable seleccionada para X
                [indx_Y,tf_Y1] = listdlg('PromptString',{'Seleccione la variable para el eje Y (vertical) [↓]','i.e : Torque'},'ListString',histogram_columnNames,'SelectionMode','single','ListSize',[350 550]);
            end
            histogram_XY(Files_data,i,indx_X,indx_Y,excel_hist)
        end
    end
catch
cd(root);
end
cd(root);

%%
function h = histogram_XY(Files_data,i,indx_X,indx_Y,excel_hist)
global Summary_filename sheetname X_min_edge X_max_edge Y_min_edge Y_max_edge XYbins; %#ok<GVMIS>
        % Obtener los nombres de las columnas
        histogram_columnNames = Files_data{i}.Properties.VariableNames;
        % Aplicar la selección de columnas para X, Y 
        Xdata = table2array(Files_data{i}(:,indx_X));
        Xname = histogram_columnNames(indx_X);
        Ydata = table2array(Files_data{i}(:,indx_Y));
        Yname = histogram_columnNames(indx_Y);
        if i == 1     % Crear una ventana emergente para introducir los valores en la primera iteracion
            prompt = {sprintf('Min X: %s',Xname{:}), sprintf('Max X: %s',Xname{:}),...
                      sprintf('Min Y: %s',Yname{:}), sprintf('Max Y: %s',Yname{:}), 'bin dimention [X Y] = [→ ↓]:'};
            dlgtitle = 'Input for Histogram Data';
            dims = [1 50];
            definput = {'0', '2970','-1000', '4600', '[45 80]'};
            inputValues = inputdlg(prompt, dlgtitle, dims, definput);
            % Asignar los valores introducidos a las variables correspondientes
            X_min_edge = str2double(inputValues{1});
            X_max_edge = str2double(inputValues{2});
            Y_min_edge = str2double(inputValues{3});
            Y_max_edge = str2double(inputValues{4});
            XYbins = str2num(inputValues{5}); %#ok<*ST2NM>
        end
        % X_edges = round(X_min_edge:((X_max_edge-(X_min_edge))/(XYbins(1))):X_max_edge,0); % Manually defined Torque edges
        % Y_edges = round(Y_min_edge:((Y_max_edge-(Y_min_edge))/(XYbins(2))):Y_max_edge,0); % Manually defined Rpm edges
        FDR=
        X_edges = [0,228,456,683,911,1139,1367,1595,1822,2050,2278,2506,2583];
        Y_edges = [-2275,-2051,-1827,-1603,-1379,-1155,-931,-707,-483,-371,-259,-203,-147,-91,-35,35,91,147,203,259,371,483,707,931,1155,1379,1603,1827,2051,2275]; % Manually defined Rpm edges
        
        figure('Name',sprintf('%s',sheetname))
        h = histogram2(Xdata,Ydata,X_edges,Y_edges,'DisplayStyle', 'bar3', 'FaceColor', 'flat');
        xlabel(Xname,'Interpreter','none');
        ylabel(Yname,'Interpreter', 'none');
        view([0 0 -1]);                                                      % [a, b] = view([0 0 -1]);
        if excel_hist ==1
            % exportamos las matrices a excel
            RLDA = (h.Values)';
            excel_export_duty(RLDA,Y_edges,X_edges,Summary_filename,sheetname)
        end    
end
function excel_export_duty(RLDA,Y_edges,X_edges,Summary_filename,sheetname)%exportamos las matrices a excel
    Y_min_bin = (Y_edges(1:end-1))';
    Y_max_bin = (Y_edges(2:end))';
    X_min_bin = (X_edges(1:end-1)); 
    X_max_bin = (X_edges(2:end));  
    try
    writematrix(Xdata,Summary_filename,"Sheet",sprintf('hist_%s',sheetname),'Range','A2','PreserveFormat',true,'AutoFitWidth',0,'FileType','spreadsheet');
    writematrix(Ydata,Summary_filename,"Sheet",sprintf('hist_%s',sheetname),'Range','B2','PreserveFormat',true,'AutoFitWidth',0,'FileType','spreadsheet');
    catch
    disp('The data block starting at cell AD4/AE4/AF4 exceeds the sheet boundaries by [-] row(s) and [-] column(s).')
    end
    % BIN DATA
    writematrix(X_min_bin,Summary_filename,"Sheet",sprintf('hist_%s',sheetname),'Range','E2','PreserveFormat',true,'AutoFitWidth',0,'FileType','spreadsheet');
    writematrix(X_max_bin,Summary_filename,"Sheet",sprintf('hist_%s',sheetname),'Range','E3','PreserveFormat',true,'AutoFitWidth',0,'FileType','spreadsheet');
    writematrix(Y_min_bin,Summary_filename,"Sheet",sprintf('hist_%s',sheetname),'Range','C4','PreserveFormat',true,'AutoFitWidth',0,'FileType','spreadsheet');
    writematrix(Y_max_bin,Summary_filename,"Sheet",sprintf('hist_%s',sheetname),'Range','D4','PreserveFormat',true,'AutoFitWidth',0,'FileType','spreadsheet');
    writematrix(RLDA,Summary_filename,"Sheet",sprintf('hist_%s',sheetname),'Range','E4','PreserveFormat',true,'AutoFitWidth',0,'FileType','spreadsheet');
   end




