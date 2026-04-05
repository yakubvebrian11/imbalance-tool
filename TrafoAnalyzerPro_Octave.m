function beban_trafo_gui
  cover = figure('Position',[400 200 600 400], ...
               'Name','Three-Phase Transformer Load Analysis', ...
               'NumberTitle','off', 'MenuBar','none', ...
               'Color',[0.95 0.95 0.95]);

  uicontrol('Style','text','Parent',cover, ...
            'String','Three-Phase Transformer Load Analysis', ...
            'FontSize',18,'FontWeight','bold', ...
            'Position',[50 250 500 60], ...
            'BackgroundColor',[0.95 0.95 0.95]);

  uicontrol('Style','text','Parent',cover, ...
            'String',sprintf('Yakub Vebrian\n2110501022\nTidar University'), ...
            'FontSize',9, ...
            'HorizontalAlignment','center', ...
            'Position',[100 200 400 40], ...
            'BackgroundColor',[0.95 0.95 0.95]);

  uicontrol('Style','pushbutton','String','Start Analysis →', ...
            'FontSize',12,'FontWeight','bold', ...
            'BackgroundColor',[0.2 0.6 0.9],'ForegroundColor','white', ...
            'Position',[220 120 160 50], ...
            'Callback',@(src,evt) mulai_program());

  uicontrol('Style','text','Parent',cover, ...
            'String','© 2025 | Transformer Load Analysis Program - Yakub Vebrian', ...
            'FontSize',6, ...
            'Position',[120 30 360 20], ...
            'BackgroundColor',[0.95 0.95 0.95], ...
            'ForegroundColor',[0.4 0.4 0.4]);

  function mulai_program()
      close(cover);
      beban_trafo_main();
  end
end

% GUI utama
function beban_trafo_main
   fig = figure('Position',[200 140 1000 620], ...
               'Name','Three-Phase Transformer Load Analysis - Multi Data', ...
               'NumberTitle','off', 'MenuBar','none');

  input_panel = uipanel('Parent',fig,'Title','Input Data (VA, IR, IS, IT, RN, Vphase)', ...
                        'FontWeight','bold','Position',[0.02 0.40 0.96 0.58]);

  kolom = {'VA','IR','IS','IT','RN','Vphase'};
  data_table = uitable('Parent',input_panel,'Data',cell(6,6), ...
                       'ColumnName',kolom,'ColumnEditable',true(1,6), ...
                       'Position',[20 60 940 270]);

  %VALIDASI
  set(data_table, 'CellEditCallback', @(src, evt) validasi_input(src, evt));

  function validasi_input(src, evt)
    newData = evt.NewData;
    if ischar(newData)
      newDataNum = str2double(newData);
    else
      newDataNum = newData;
    end

    % pringatan
    if isnan(newDataNum)
      msgbox('Input harus berupa angka!','Peringatan','warn');
      d = get(src,'Data');
      d{evt.Indices(1), evt.Indices(2)} = [];
      set(src,'Data',d);
    end
  end

  % TOMBOL-TOMBOL
  uicontrol(input_panel,'Style','pushbutton','String','+ Row', ...
            'Position',[20 20 90 30], 'Callback',@(s,e) tambah_baris());
  uicontrol(input_panel,'Style','pushbutton','String','- Row', ...
            'Position',[120 20 90 30], 'Callback',@(s,e) hapus_baris());
  uicontrol(input_panel,'Style','pushbutton','String','Import Excel', ...
            'Position',[220 20 110 30], 'Callback',@(s,e) impor_excel());
  uicontrol(input_panel,'Style','pushbutton','String','Run Analysis', ...
            'Position',[350 20 140 30], 'BackgroundColor',[0.2 0.6 0.9], ...
            'ForegroundColor','white','FontWeight','bold', 'Callback',@(s,e) hitung());
  uicontrol(input_panel,'Style','pushbutton','String','Simpan Hasil', ...
            'Position',[505 20 140 30], 'BackgroundColor',[0.3 0.7 0.3], ...
            'ForegroundColor','white','FontWeight','bold', 'Callback',@(s,e) simpan_hasil());
  uicontrol(input_panel,'Style','pushbutton','String','Export Excel', ...
            'Position',[660 20 180 30], 'BackgroundColor',[0.65 0.55 0.25], ...
            'ForegroundColor','white','FontWeight','bold', 'Callback',@(s,e) export_excel());

   % Tombol Bersihkan Input (mengosongkan isi tabel)
   uicontrol(input_panel,'Style','pushbutton','String','Clear Input',...
          'Position',[860 20 140 30], 'Callback',@(s,e) bersihkan_input());

   % Tombol Reset GUI (mengembalikan GUI seperti awal)
   uicontrol(input_panel,'Style','pushbutton','String','Reset Result',...
          'Position',[1010 20 120 30], 'Callback',@(s,e) reset_hasil());

  % PANEL OUTPUT
  output_panel = uipanel('Parent',fig,'Title','Analysis Result', ...
                         'FontWeight','bold','Position',[0.02 0.02 0.96 0.36]);

  output_box = uicontrol('Parent',output_panel,'Style','edit','Max',200, ...
                         'HorizontalAlignment','left','Position',[20 20 940 170], ...
                         'Enable','inactive','FontName','FixedWidth','FontSize',10);
  hasil_terakhir_text = '';
  hasil_terakhir_table = {};
  hasil_header = {'Idx','VA','IR','IS','IT','RN','Vphase','IN', ...
                  'Ibeban penuh(A)','Irata-rata (A)','Load(%)','Kond_Load', ...
                  'CU(%)','Kond_CU','P_loss_N(kW)', ...
                  'Rugi_bln(Rp)'};
  sim_log_lines = {};
  sim_CU_iter = [];
  sim_arus_iter = [];
  sim_idx_data = NaN;

  % ====== CALLBACKS ======

  function tambah_baris()
    d = get(data_table,'Data');
    d(end+1,:) = cell(1,6);
    set(data_table,'Data',d);
  end

  function hapus_baris()
    d = get(data_table,'Data');
    if size(d,1) > 1
      d(end,:) = [];
      set(data_table,'Data',d);
    end
  end

  function bersihkan_input()
    data = get(data_table, 'Data');
    [r,c] = size(data);
    data(:,:) = {[]};
    set(data_table,'Data',data);
  end

  function reset_hasil()
    try
        set(output_box, 'String', '');
    catch
    end
    try
        figs = findall(0,'Type','figure');
        for f = figs'
            if f ~= fig
                close(f);
            end
        end
    catch
    end
    hasil_terakhir_text  = '';
    hasil_terakhir_table = {};
    sim_log_lines = {};
    sim_CU_iter = [];
    sim_arus_iter = [];
    sim_idx_data = NaN;
    msgbox('GUI berhasil di-reset.','Sukses');
  end

  function impor_excel()
    [file,path] = uigetfile({'*.xls;*.xlsx'},'Pilih File Excel (6 kolom)');
    if isequal(file,0), return; end
    fullfile_name = fullfile(path,file);
    try
      try pkg load io; catch; end
      [num,txt,raw] = xlsread(fullfile_name);
      if ~isempty(raw) && any(cellfun(@ischar, raw(1,:)))
        raw(1,:) = [];
      end
      if size(raw,2) >= 6
        raw = raw(:,1:6);
      else
        error('Kolom kurang dari 6.');
      end
      set(data_table,'Data',raw);
      msgbox('Import berhasil. Cek tabel input.','Sukses');
    catch ME
      msgbox(['Gagal impor: ' ME.message],'Error','error');
    end
    end

  function hitung()
    data = get(data_table,'Data');
    hasil_txt = '';
    hasil_tbl = {};
    CU_list = nan(size(data,1),1);

    for i = 1:size(data,1)
      row = data(i,:);
      vals = zeros(1,6);
      for j = 1:6
        v = row{j};
        if isempty(v)
          vals(j) = NaN;
        elseif isnumeric(v)
          vals(j) = v;
        else
          vals(j) = str2double(num2str(v));
        end
      end
      if any(isnan(vals))
        hasil_txt = [hasil_txt, sprintf('Baris %d: Data tidak lengkap atau bukan angka valid!\n', i)];
        continue;
      end

      % Ekstrak
      KVA = vals(1); IR = vals(2); IS = vals(3); IT = vals(4);
      RN = vals(5); V_phase = vals(6);

      % --- Hitung Arus Netral (IN) otomatis ---
      IRx = IR*cosd(0);    IRy = IR*sind(0);
      ISx = IS*cosd(-120); ISy = IS*sind(-120);
      ITx = IT*cosd(120);  ITy = IT*sind(120);
      INx = IRx + ISx + ITx;
      INy = IRy + ISy + ITy;
      IN = sqrt(INx^2 + INy^2);

      % --- Perhitungan beban dan ketidakseimbangan ---
      I_F = KVA / (V_phase*sqrt(3));
      I_avg = mean([IR IS IT]);
      Load_percent = (I_avg/I_F)*100;
      kondisi_beban = 'Buruk';
      if Load_percent < 80, kondisi_beban = 'Baik'; end

      CU = ((abs(IR/I_avg-1) + abs(IS/I_avg-1) + abs(IT/I_avg-1))/3)*100;
      CU_list(i) = CU;
      if CU<10
        kondisi_CU='Good';
      elseif CU<20
        kondisi_CU='Fair';
      elseif CU<=25
        kondisi_CU='Poor';
      else
        kondisi_CU='Sereve';
      end

      % --- Rugi daya dan finansial ---
      P_loss_N = (IN^2 * RN)/1000;
      Tarif = 1444.70;
      Rugi  = P_loss_N * 24 * 4 * Tarif;

      hasil_txt = [hasil_txt, sprintf([ ...
        '=== Date-%d ===\n' ...
        'Full-load Current             : %.2f A\n' ...
        'Average Current              : %.2f A\n' ...
        'Neutral Current               : %.2f A\n' ...
        'Load Percentage              : %.2f %% (%s)\n' ...
        'Load Unbalance Percentage : %.2f %% (%s)\n' ...
        'Power Loss : %.4f kW\n' ...
        'Finansial Loss                : Rp %.0f\n\n'], ...
        i, I_F, I_avg, IN, Load_percent, kondisi_beban, ...
        CU, kondisi_CU, P_loss_N, Rugi)];


      hasil_tbl(end+1,:) = {i, KVA, IR, IS, IT, RN, V_phase, IN, ...
                            I_F, I_avg, Load_percent, kondisi_beban, ...
                            CU, kondisi_CU, P_loss_N, Rugi};
    end


    % --- Simulasi untuk CU tertinggi ---

    sim_log_lines = {}; sim_CU_iter = []; sim_arus_iter = []; sim_idx_data = NaN;
    [max_CU, idx_maxCU] = max(CU_list);
    if ~isnan(max_CU) && max_CU > 10
      sim_idx_data = idx_maxCU;
      row = data(idx_maxCU,:);
      vals = zeros(1,6);
      for j=1:6
        v=row{j};
        if isempty(v), vals(j)=NaN;
        elseif isnumeric(v), vals(j)=v;
        else vals(j)=str2double(num2str(v));
        end
      end

      IR=vals(2); IS=vals(3); IT=vals(4);
      I_avg = mean([IR IS IT]);

      arus = [IR, IS, IT];
      iter = 0;
      pindah_matrix = zeros(3,3);
      fasa_label = {'R','S','T'};
      CU_iter = [];
      AR_iter = [];

      while true
        iter = iter + 1;

        % cari fasa terbesar & terkecil
        [max_val, idx_max] = max(arus);
        [min_val, idx_min] = min(arus);
        fasa_max = fasa_label{idx_max};
        fasa_min = fasa_label{idx_min};

        % hitung arus yang dipindahkan (α = 0.05)
        delta_I = 0.1 * (max_val - min_val);

        % perbarui arus
        arus(idx_max) = arus(idx_max) - delta_I;
        arus(idx_min) = arus(idx_min) + delta_I;

        pindah_matrix(idx_max, idx_min) = pindah_matrix(idx_max, idx_min) + delta_I;

        % hitung CU baru
        CU_baru = ((abs(arus(1)/I_avg-1) + abs(arus(2)/I_avg-1) + abs(arus(3)/I_avg-1))/3)*100;
        CU_iter(end+1) = CU_baru;
        AR_iter(:,end+1) = arus(:);

        % hasil per iterasi
        log_line = sprintf(['Iterasi %d: \n' ...
          '  Fasa tertinggi  : %s (%.2f A)\n' ...
          '  Fasa terendah   : %s (%.2f A)\n' ...
          '  ΔI (2%% × Δ)     : %.2f A dipindah dari %s → %s\n' ...
          '  Arus baru (R/S/T): %.2f / %.2f / %.2f A\n' ...
          '  CU baru          : %.2f %%\n\n'], ...
          iter, fasa_max, max_val, fasa_min, min_val, delta_I, fasa_max, fasa_min, ...
          arus(1), arus(2), arus(3), CU_baru);

        sim_log_lines{end+1,1} = log_line;

        if CU_baru <= 1
          break;
        end
      end


      % ringkasan perpindahan
      ringkasan_pindah = '';
      for ii = 1:3
        for jj = 1:3
          if ii~=jj && pindah_matrix(ii,jj) > 0
            ringkasan_pindah = [ringkasan_pindah, ...
              sprintf('Total pindah dari Fasa-%s ke Fasa-%s: %.2f A\n', ...
              fasa_label{ii}, fasa_label{jj}, pindah_matrix(ii,jj))];
          end
        end
      end

      hasil_txt = [hasil_txt, sprintf([ ...
        '>> Simulasi Penyeimbangan Iteratif (Data ke-%d; CU awal %.2f%%) <<\n'], ...
        idx_maxCU, max_CU)];
      for k = 1:numel(sim_log_lines)
      hasil_txt = [hasil_txt, sprintf('%s\n', sim_log_lines{k})];
    end
      hasil_txt = [hasil_txt, sprintf('CU Akhir: %.2f %%\n', CU_iter(end))];
      hasil_txt = [hasil_txt, ringkasan_pindah];

      % simpan utk ekspor/plot
      sim_CU_iter = CU_iter;
      sim_arus_iter = AR_iter;

      % grafik
      figure('Name','Grafik Simulasi Penyeimbangan (CU tertinggi)');
      subplot(2,1,1);
      plot(1:length(CU_iter), CU_iter, '-o','LineWidth',2);
      xlabel('Iterasi'); ylabel('CU (%)'); title('Perubahan CU per Iterasi'); grid on;

      subplot(2,1,2);
      plot(1:length(CU_iter), AR_iter(1,:),' -o','LineWidth',2); hold on;
      plot(1:length(CU_iter), AR_iter(2,:),' -o','LineWidth',2);
      plot(1:length(CU_iter), AR_iter(3,:),' -o','LineWidth',2);
      xlabel('Iterasi'); ylabel('Arus (A)'); legend('R','S','T','Location','best');
      title('Perubahan Arus R/S/T per Iterasi'); grid on;
    end

    set(output_box,'String',hasil_txt);
    hasil_terakhir_text  = hasil_txt;
    hasil_terakhir_table = [hasil_header; hasil_tbl];
  end


  function simpan_hasil()
    if isempty(hasil_terakhir_text)
      msgbox('Belum ada hasil. Klik "Hitung Analisis" dulu.','Info'); return;
    end
    [file,path] = uiputfile({'*.txt';'*.xlsx'},'Simpan Hasil Analisis');
    if isequal(file,0), return; end
    [~,~,ext] = fileparts(file);
    fullname = fullfile(path,file);

    try
      switch lower(ext)
        case '.txt'
          fid = fopen(fullname,'w');
          if fid==-1, error('Gagal membuka file.'); end
          fprintf(fid,'%s',hasil_terakhir_text);
          fclose(fid);
        case '.xlsx'
          try, pkg load io; catch, end
          % sheet 1: hasil ringkas per baris
          xlswrite(fullname, hasil_terakhir_table, 'Hasil');
          % sheet 2: log simulasi (kalau ada)
          if ~isempty(sim_log_lines)
            xlswrite(fullname, [{'Log Simulasi'}; sim_log_lines], 'Simulasi_Log');
          end
          % sheet 3: data iterasi (CU & arus)
          if ~isempty(sim_CU_iter)
            iter_idx = num2cell((1:numel(sim_CU_iter))');
            CUcol = num2cell(sim_CU_iter(:));
            sheet3 = [{'Iter','CU(%)'}; [iter_idx, CUcol]];
            xlswrite(fullname, sheet3, 'Sim_CU');

            if ~isempty(sim_arus_iter)
              sheet4 = [{'Iter','I_R','I_S','I_T'}];
              for k=1:size(sim_arus_iter,2)
                sheet4(end+1,:) = {k, sim_arus_iter(1,k), sim_arus_iter(2,k), sim_arus_iter(3,k)}; %#ok<AGROW>
              end
              xlswrite(fullname, sheet4, 'Sim_Arus');
            end
          end
        otherwise
          error('Ekstensi tidak didukung.');
      end
      msgbox('Hasil berhasil disimpan.','Sukses');
    catch ME
      msgbox(['Gagal simpan: ' ME.message],'Error','error');
    end
  end

  function export_excel()
    % Export INPUT + HASIL ringkas ke satu file Excel
    if isempty(hasil_terakhir_text)
      msgbox('Belum ada hasil. Klik "Hitung Analisis" dulu.','Info'); return;
    end
    d = get(data_table,'Data');
    if isempty(d)
      msgbox('Tabel input kosong.','Info'); return;
    end
    [file,path] = uiputfile('Export_Input_Hasil.xlsx','Nama file export');
    if isequal(file,0), return; end
    fullname = fullfile(path,file);
    try
      try, pkg load io; catch, end
      % Sheet Input
      xlswrite(fullname, [kolom; d], 'Input_Data');
      % Sheet Hasil kalau sudah ada
      if ~isempty(hasil_terakhir_table)
        xlswrite(fullname, hasil_terakhir_table, 'Hasil_Analisis');
      end
      msgbox('Export Excel selesai.','Sukses');
    catch ME
      msgbox(['Gagal export: ' ME.message],'Error','error');
    end
  end
end

