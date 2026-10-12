%% rlc_to_xlsx.m - Export the Channel Overview of an RLC file to Excel
% Requirements: System Monitor must be open with a project loaded.
% Use a copy of the project because RemoteLoggingLoad loads the RLC into it.

clear all; clc;

% ---------------- CONFIGURATION ----------------
dllPath = 'C:\Program Files (x86)\McLaren Electronic Systems\System Monitor 8\SystemMonitorAPI.dll';
% Select the RLC configuration file
[rlcFile, rlcPath] = uigetfile( ...
    {'*.rlc', 'RLC Configuration Files (*.rlc)'}, ...
    'Select an RLC configuration file');
% Stop if the user cancels the selection
if isequal(rlcFile, 0)
    disp('File selection cancelled. Export aborted.');
    return;
end
% Build the full file path
rlcFullPath = fullfile(rlcPath, rlcFile);
fprintf('Selected RLC file: %s\n', rlcFullPath);
rateDivisor = 1024; % Adjust if the reported KB/s values do not match System Monitor
% ------------------------------------------------
% Create the output folder next to this script if it does not exist
scriptDir = fileparts(mfilename('fullpath'));
if isempty(scriptDir)
    scriptDir = pwd;
end
outputDir = fullfile(scriptDir,'output');
if ~isfolder(outputDir)
    mkdir(outputDir);
end
xlsxPath = fullfile(outputDir,['RLC_export_' char(datetime('now','Format','yyyyMMdd_HHmmss')) '.xlsx']);
NET.addAssembly(dllPath);
api = SystemMonitorAPI.TSysMonApiClass;

% 1) Load the RLC file
rc = api.RemoteLoggingLoad(rlcFullPath);
assert(double(rc)==0,'RemoteLoggingLoad returned %d',double(rc));

% 2) Read remote parameter details
% Outputs: rc, appIds, paramIds, names, descriptions, channel names,
% sizes, values, log types, parameter count, channel count, slots.
eo = NET.createArray('System.Object',0);
o = tryCall(12,...
    @() api.GetRemoteParamDetails([],[],[],[],[],[],[],[],[]),...
    @() api.GetRemoteParamDetails(eo,eo,eo,eo,eo,eo,eo,eo,eo));
fprintf('Output classes: appIds=%s ids=%s vals=%s\n',class(o{2}),class(o{3}),class(o{8}));
appIds = toNum(o{2});
appIds = appIds(:);
pIds = toStr(o{3});
pNames = toStr(o{4});
pDesc = toStr(o{5});
chNames = toStr(o{6});
pSize = toNum(o{7});
pSize = pSize(:);
V = toNum(o{8});
nP = double(o{10});
nC = double(o{11});
fprintf('nP=%d nC=%d size(V)=%s\n',nP,nC,mat2str(size(V)));
% Arrange the values matrix as parameters-by-channels
if isequal(size(V),[nP nC])
    F = V;
elseif isequal(size(V),[nC nP])
    F = V.';
elseif numel(V)==nP*nC
    F = reshape(V,[nC nP]).';
    warning('The values array was flattened; verify the first row.');
else
    error('Unexpected values dimensions: %s',mat2str(size(V)));
end
% Use the parameter ID as the displayed name; fall back to the name
pShow = pIds;
pShow(pShow=="") = pNames(pShow=="");
fprintf('Row 1 (%s): %s\n',pShow(1),mat2str(F(1,:)));

% 3) Resolve application names
appName = "APP"+string(appIds);
try
    oa = tryCall(3,@() api.GetAppDetails([],[]));
    ids = toNum(oa{2});
    ids = ids(:);
    names = toStr(oa{3});
    for i = 1:nP
        k = find(ids==appIds(i),1);
        if ~isempty(k)
            appName(i) = names(k);
        end
    end
catch ME
    warning('GetAppDetails failed: %s',ME.message);
end

% 4) Read the Prime flag
% The method signatures must match the installed API DLL.
prime = strings(nP,1);
try
    for i = 1:nP
        a = int16(appIds(i));
        id = char(pShow(i));
        op = tryCall(10,...
            @() api.GetParameterPropertiesEx(a,id,[]),...
            @() api.GetParameterPropertiesEx(a,id));
        if logical(op{8})
            prime(i) = "X";
        end
    end
catch ME
    warning('GetParameterPropertiesEx unavailable; Prime values may be blank: %s',ME.message);
end

% 5) Read channel type flags
% L indicates logging; T indicates telemetry.
chType = strings(1,nC);
try
    for c = 1:nC
        ch = int16(c-1);
        oc = tryCall(6,...
            @() api.GetRemoteChannelPropertiesAll(ch),...
            @() api.GetRemoteChannelPropertiesAll(ch,'',false,false,false,int16(0)));
        t = "";
        if logical(oc{3})
            t = "L";
        end
        if logical(oc{4})
            t = t+"T";
        end
        chType(c) = t;
    end
catch ME
    warning('GetRemoteChannelPropertiesAll unavailable; channel types may be blank: %s',ME.message);
end

% 6) Read the logging rates
kbs = strings(1,nC);
try
    orr = tryCall(3,@() api.GetRemoteLoggingRates([],[]));
    rLog = toNum(orr{2});
    rLog = rLog(:);
    rTel = toNum(orr{3});
    rTel = rTel(:);
    fprintf('Rates: logging=%d telemetry=%d channels=%d\n',numel(rLog),numel(rTel),nC);
    for c = 1:nC
        if chType(c)=="T" && numel(rTel)>=c
            r = rTel(c);
        elseif numel(rLog)>=c
            r = rLog(c);
        else
            continue;
        end
        kbs(c) = sprintf('%.1f KB/S',r/rateDivisor);
    end
catch ME
    warning('GetRemoteLoggingRates failed: %s',ME.message);
end

% 7) Build the worksheet data
hdr1 = [{'Parameter','Description','Application','Bytes','Prime'},cellstr(chNames(:).'),{'Note','Reason for logging'}];
hdr2 = [repmat({''},1,5),cellstr(chType),{'',''}];
hdr3 = [repmat({''},1,5),cellstr(kbs),{'',''}];
body = cell(nP,5+nC+2);
for i = 1:nP
    body(i,1:5) = {char(pShow(i)),char(pDesc(i)),char(appName(i)),pSize(i),char(prime(i))};
    for c = 1:nC
        if F(i,c)>0
            body{i,5+c} = F(i,c);
        else
            body{i,5+c} = '---';
        end
    end
    body(i,end-1:end) = {'',''};
end
allData = [hdr1;hdr2;hdr3;body];
assert(size(allData,2)==5+nC+2,'Excel column count mismatch.');

% 8) Write the workbook
writecell(allData,xlsxPath,'Sheet','Channel Overview');
assert(isfile(xlsxPath),'Excel export failed: output file was not created.');
fprintf('Exported successfully: %s\n',xlsxPath);

% 9) Format the workbook using Excel COM automation
% Formatting is optional; the exported workbook remains usable if formatting fails.
try
    xl = actxserver('Excel.Application');
    xl.Visible = false;
    xl.DisplayAlerts = false;
    wb = xl.Workbooks.Open(xlsxPath);
    ws = wb.Worksheets.Item(1);
    lastRow = size(allData,1);
    lastColNum = size(allData,2);
    lastCol = excelCol(lastColNum);
    lastChannelCol = excelCol(5+nC);
    noteColNum = 5+nC+1;
    noteCol = excelCol(noteColNum);
    % Format the header rows
    ws.Range(sprintf('A1:%s3',lastCol)).Font.Bold = true;
    ws.Range(sprintf('A1:%s3',lastCol)).VerticalAlignment = -4108;
    ws.Rows.Item(1).RowHeight = 24;
    ws.Rows.Item(2).RowHeight = 20;
    ws.Rows.Item(3).RowHeight = 20;
    % Set explicit widths for the descriptive columns
    ws.Columns.Item(1).ColumnWidth = 28;
    ws.Columns.Item(2).ColumnWidth = 35;
    ws.Columns.Item(3).ColumnWidth = 24;
    ws.Columns.Item(4).ColumnWidth = 10;
    ws.Columns.Item(5).ColumnWidth = 10;
    % Set channel column widths
    if nC>0
        ws.Range(sprintf('F1:%s%d',lastChannelCol,lastRow)).ColumnWidth = 12;
    end
    ws.Columns.Item(noteColNum).ColumnWidth = 40;
    ws.Columns.Item(noteColNum+1).ColumnWidth = 40;
    % Format channel values as frequencies in Hz
    if nC>0 && nP>0
        dataRange = ws.Range(sprintf('F4:%s%d',lastChannelCol,lastRow));
        dataRange.NumberFormat = '0 "Hz"';
        dataRange.HorizontalAlignment = -4108;
        dataRange.VerticalAlignment = -4108;
    end
    % Align parameter and application details
    ws.Range(sprintf('A4:C%d',lastRow)).HorizontalAlignment = -4131;
    ws.Range(sprintf('D4:E%d',lastRow)).HorizontalAlignment = -4108;
    % Highlight the entire parameter row when its Note cell is not blank
    % The conditional formatting remains active when the workbook is edited in Excel.
    dataRows = ws.Range(sprintf('A4:%s%d',lastCol,lastRow));
    formula = sprintf('=$%s4<>""',noteCol);
    formatConditions = dataRows.FormatConditions;
    yellowRule = formatConditions.Add(2,[],formula);
    yellowRule.Interior.Color = 65535;
    % Enable filtering using the third row as the column header
    ws.Range(sprintf('A3:%s%d',lastCol,lastRow)).AutoFilter;
    % Freeze the top three rows and the first column
    ws.Activate;
    xl.ActiveWindow.SplitRow = 3;
    xl.ActiveWindow.SplitColumn = 1;
    xl.ActiveWindow.FreezePanes = true;
    % Save and close Excel
    wb.Save;
    wb.Close(false);
    xl.Quit;
    delete(xl);
    fprintf('Excel formatting and conditional highlighting completed.\n');
catch ME
    warning('Excel formatting failed; the exported workbook may still be valid: %s',ME.message);
    try
        wb.Close(false);
    catch
    end
    try
        xl.Quit;
        delete(xl);
    catch
    end
end

%% Local helper functions
function out = tryCall(nOut,varargin)
% Try each API signature until one succeeds and returns rc == 0.
% The first output is the return code; the remaining outputs contain the data.
lastErr = MException('tryCall:none','No working API signature found');
for k = 1:numel(varargin)
    try
        out = cell(1,nOut);
        [out{1:nOut}] = varargin{k}();
        assert(double(out{1})==0,'Return code %d',double(out{1}));
        return;
    catch ME
        lastErr = ME;
        fprintf(2,'  tryCall variant %d failed: %s\n',k,strtrim(ME.message));
    end
end
rethrow(lastErr);
end
function n = toNum(x)
% Convert .NET arrays and SafeArrays to MATLAB doubles.
try
    n = double(x);
catch
    c = cell(x);
    n = cellfun(@double,c);
end
end
function s = toStr(x)
% Convert .NET string arrays to MATLAB string arrays.
if isstring(x) || ischar(x)
    s = string(x);
    s = s(:);
    return;
end
c = cell(x);
s = strings(numel(c),1);
for i = 1:numel(c)
    s(i) = string(char(c{i}));
end
end
function col = excelCol(n)
% Convert a positive integer to an Excel column label.
assert(n>=1 && n==floor(n),'Excel column index must be a positive integer.');
col = '';
while n>0
    r = mod(n-1,26);
    col = [char(65+r),col];
    n = floor((n-1)/26);
end
end
