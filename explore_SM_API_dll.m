% --- CONFIG ---
smDir  = 'C:\Program Files (x86)\McLaren Electronic Systems\System Monitor 8'; %Folder containing SystemMonitorAPI.dll
outDir = fullfile(pwd, 'sm_dump');
% --------------
assert(isfolder(smDir), 'Folder not found: %s', smDir);
if ~isfolder(outDir), mkdir(outDir); end
dlls = dir(fullfile(smDir, '**', '*.dll'));
fprintf('DLLs found: %d\n', numel(dlls));
flagType = System.Type.GetType('System.Reflection.BindingFlags');
flags = System.Enum.Parse(flagType, 'Public,Instance,Static,DeclaredOnly');
skipped = strings(0,1);
for k = 1:numel(dlls)
    p = fullfile(dlls(k).folder, dlls(k).name);
    try
        a = System.Reflection.Assembly.LoadFrom(p);   % Load assembly for reflection
        NET.addAssembly(p);                           % Load assembly for use in MATLAB
        types = a.GetExportedTypes();
    catch ME
        skipped(end+1,1) = string(dlls(k).name) + "  ->  " + string(ME.message); %#ok<SAGROW>
        continue
    end
    fid = fopen(fullfile(outDir, [dlls(k).name '.txt']), 'w');
    for t = 1:types.Length
        ty = types(t);
        fprintf(fid, '\n=== %s ===\n', char(ty.FullName));
        try
            try, props = ty.GetProperties(flags); catch, props = ty.GetProperties(); end
            for i = 1:props.Length
                fprintf(fid, '  [P] %s %s\n', char(props(i).PropertyType.Name), char(props(i).Name));
            end
            try, meths = ty.GetMethods(flags); catch, meths = ty.GetMethods(); end
            for i = 1:meths.Length
                m = meths(i);
                if m.IsSpecialName, continue; end
                ps = m.GetParameters();
                sig = strings(1, ps.Length);
                for j = 1:ps.Length
                    sig(j) = string(char(ps(j).ParameterType.Name)) + " " + string(char(ps(j).Name));
                end
                fprintf(fid, '  [M] %s %s(%s)\n', char(m.ReturnType.Name), char(m.Name), strjoin(sig, ', '));
            end
        catch
        end
    end
    fclose(fid);
    fprintf('OK  %s (%d types)\n', dlls(k).name, types.Length);
end
fprintf('\nSkipped assemblies (native or missing dependencies): %d\n', numel(skipped));
disp(skipped)
% Search for keywords
kw = ["rlc","RemoteLog","Logging","Telem","Trigger","Slot","Frequency","Prime","PPO"];
files = dir(fullfile(outDir, '*.txt'));
for k = 1:numel(files)
    txt = fileread(fullfile(files(k).folder, files(k).name));
    hits = kw(arrayfun(@(w) contains(txt, w, 'IgnoreCase', true), kw));
    if ~isempty(hits)
        fprintf('%-40s -> %s\n', files(k).name, strjoin(hits, ', '));
    end
end