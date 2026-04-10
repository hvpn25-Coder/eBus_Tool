%% Model Info (very compact)

MS = sMP.cfg.Configuration.ModuleSetup;
nM = numel(MS); nF = 6;
nDS = max(arrayfun(@(x)numel(x.DataSet),MS)) ;

FN = strings(nF,nM); FV = FN;
DVN = cell (nDS,nM) ; DCN = DVN;

for m=l:nM
    f = string (fieldnames (MS (m) .Module)); v = string(struct2cell (MS (m) .Module));
    k = 1:min(nF,numel(f)); FN(k,m)=f(k); FV(k,m)=v(k);
    ds = MS (m).DataSet;
    if ~isempty(ds)
       DVN(1:numel(ds),m) = {ds.variant};
       DCN(1:numel(ds),m) = {ds.className};
    end 
end 

CAT = FV(2,:);
DN = strings(size(DVN)); DN(cellfun(@(x)ischar(x)||isstring(x),DVN)) = string(DVN(cellfun(@(x)ischar(x)||isstring(x),DVN)));
DV = strings(size(DCN)); DV(cellfun(@(x)ischar(x)||isstring(x),DCN)) = string(DCN(cellfun(@(x)ischar(x)||isstring(x),DCN)));

DIVe_Mod_Data = struct; 

for k=1:nM
    c = CAT (k);
    DIVe_Mod_Data.(c+"_a") = join(FN(:,k),newline);
    DIVe_Mod_Data.(c+"_b") = join(FV(:,k),newline) ;
    DIVe_Mod_Data.(c+"_c") = join(DV(DV(:,k)~="",k),newline);
    DIVe_Mod_Data.(c+"_d") = join(DN(DN(:,k)~="",k),newline);
end

%% Ordering
ord = ["bdry","human","phys","ctrl"];
[~,g] = ismember(FV(1,:),ord);
idx = find(g); [~,o]=sort (g(idx)); idx=idx(o);

FN=FN(:,idx); FV=FV(:,idx); DCN=DCN(:,idx); DVN=DVN(:,idx);

%% DIVe export (single generic block)

blk = { "DB","bdry",NaN
        "DH","human",NaN
        "DP","phys",10
        "DC","ctrl",20 };

for b = 1:size (blk, 1)
    A = FV(2:4, FV(1,:)==blk{b,2});
    if isempty(A), continue, end
    
    A = string (A);
    maxI = blk{b,3};

    if ~isnan(maxI) && size(A,2) < maxI
        A (:,end+1:maxI) = " ";
    end

    for i = 1:size (A,2)
        for j = 1:size(A,1)
            DIVe_Mod_Data.(blk{b,1}+"_"+j+"_"+i) = A(j,i);
        end
    end
end

