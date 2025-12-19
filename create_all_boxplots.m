
config.filename = 'data.xls';
config.num_groups = 7;
config.columns_start = 6;  % Column F
config.columns_end = 33;   % Column AG
config.save_figures = true;
config.show_figures = true;
config.output_dir = 'boxplots_output';
config.figure_format = 'png';  % 'png', 'fig', 'pdf', or 'eps'
config.show_mean = false;
config.show_median = false;
config.colors = 'b';  % Box color

%% Create output directory 
if config.save_figures && ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

%% Read the Excel file
fprintf('Reading data from %s...\n', config.filename);

% Read all four header rows
header_row0 = readcell(config.filename, 'Range', 'A1:AG1');  % Main category
header_row1 = readcell(config.filename, 'Range', 'A2:AG2');  % Units
header_row2 = readcell(config.filename, 'Range', 'A3:AG3');  % Walking type
header_row3 = readcell(config.filename, 'Range', 'A4:AG4');  % Side (Left/Right)

% Read data (starting from row 5)
data = readmatrix(config.filename, 'Range', 'A5');

% Extract group column (Skupina = column C = index 3)
all_groups = data(:, 3);

fprintf('Data loaded successfully.\n');
fprintf('Total rows: %d\n', size(data, 1));
fprintf('Creating boxplots for columns %d to %d...\n\n', ...
    config.columns_start, config.columns_end);

%% Boxplot for each column
columns_to_process = config.columns_start:config.columns_end;
num_columns = length(columns_to_process);
summary_stats = cell(num_columns, 5);  

for idx = 1:num_columns
    col_idx = columns_to_process(idx);
    
    % Get data for this column
    column_data = data(:, col_idx);
    
    % Remove NaN values
    valid_idx = ~isnan(column_data) & ~isnan(all_groups);
    valid_data = column_data(valid_idx);
    valid_groups = all_groups(valid_idx);
    
    % Skip if no valid data
    if isempty(valid_data)
        fprintf('  Column %d: No valid data, skipping...\n', col_idx);
        continue;
    end
    
    % Prepare grouped data
    group_data = [];
    group_labels = [];
    
    for g = 1:config.num_groups
        group_mask = (valid_groups == g);
        group_values = valid_data(group_mask);
        
        if ~isempty(group_values)
            group_data = [group_data; group_values];
            group_labels = [group_labels; ones(length(group_values), 1) * g];
        end
    end
    
    % Build title from headers
    title_info = build_title(header_row0, header_row1, header_row2, ...
                             header_row3, col_idx);
    
    % Create figure
    if config.show_figures
        fig = figure('Position', [100, 100, 900, 600]);
    else
        fig = figure('Visible', 'off', 'Position', [100, 100, 900, 600]);
    end
    
    % Create boxplot
    
    boxplot(group_data, group_labels, ...
    'Labels', {'1', '2', '3', '4', '5', '6', '7'});
    
    % Title and labels
    title(title_info.full_title, 'Interpreter', 'none', ...
          'FontSize', 12, 'FontWeight', 'bold');
    xlabel('Skupina (EDSS Group)', 'FontSize', 11, 'FontWeight', 'bold');
    ylabel(title_info.ylabel, 'FontSize', 11, 'FontWeight', 'bold');
    
    grid on;
    set(gca, 'FontSize', 10);
    
    % Add statistical overlays
    hold on;
    legend_entries = {};
    
    if config.show_mean
        % Collect means for all groups
        group_means = nan(1, config.num_groups);
        for g = 1:config.num_groups
            group_mask = (group_labels == g);
            if any(group_mask)
                group_means(g) = mean(group_data(group_mask));
            end
        end
        
        % Plot mean as red line
        plot(1:config.num_groups, group_means, 'r-', 'LineWidth', 2);
        legend_entries{end+1} = 'Mean';
    end
    
    if config.show_median
        % Collect medians for all groups
        group_medians = nan(1, config.num_groups);
        for g = 1:config.num_groups
            group_mask = (group_labels == g);
            if any(group_mask)
                group_medians(g) = median(group_data(group_mask));
            end
        end
        
        % Plot median as green line
        plot(1:config.num_groups, group_medians, 'g-', 'LineWidth', 2);
        legend_entries{end+1} = 'Median';
    end
    
    if ~isempty(legend_entries)
        legend(legend_entries, 'Location', 'best');
    end
    
    hold off;
    
    % Calculate summary statistics
    overall_stats = calculate_stats(group_data);
    summary_stats{idx, 1} = title_info.full_title;
    summary_stats{idx, 2} = overall_stats.mean;
    summary_stats{idx, 3} = overall_stats.median;
    summary_stats{idx, 4} = overall_stats.std;
    summary_stats{idx, 5} = overall_stats.n;
    
    % Save figure
    if config.save_figures
        safe_filename = create_safe_filename(title_info.full_title, col_idx);
        output_path = fullfile(config.output_dir, safe_filename);
        
        switch lower(config.figure_format)
            case 'png'
                saveas(fig, [output_path '.png']);
            case 'fig'
                saveas(fig, [output_path '.fig']);
            case 'pdf'
                saveas(fig, [output_path '.pdf']);
            case 'eps'
                saveas(fig, [output_path '.eps']);
            otherwise
                saveas(fig, [output_path '.png']);
        end
    end
    
    fprintf('  [%d/%d] Created: %s\n', idx, num_columns, title_info.full_title);
    
    % Close figure if not showing
    if ~config.show_figures
        close(fig);
    end
end

%% Save summary statistics
if config.save_figures
    summary_table = cell2table(summary_stats, ...
        'VariableNames', {'Variable', 'Mean', 'Median', 'StdDev', 'N'});
    writetable(summary_table, fullfile(config.output_dir, 'summary_statistics.csv'));
    fprintf('\nSummary statistics saved to: %s\n', ...
            fullfile(config.output_dir, 'summary_statistics.csv'));
end

fprintf('\n=== COMPLETE ===\n');
fprintf('Total boxplots created: %d\n', num_columns);
if config.save_figures
    fprintf('Figures saved to: %s\n', config.output_dir);
end

%% Helper Functions

function title_info = build_title(h0, h1, h2, h3, col_idx)
    % Build comprehensive title from multi-level headers
    
    % Get header values with bounds checking
    if col_idx <= length(h0)
        h0_val = h0{col_idx};
    else
        h0_val = '';
    end
    
    if col_idx <= length(h1)
        h1_val = h1{col_idx};
    else
        h1_val = '';
    end
    
    if col_idx <= length(h2)
        h2_val = h2{col_idx};
    else
        h2_val = '';
    end
    
    if col_idx <= length(h3)
        h3_val = h3{col_idx};
    else
        h3_val = '';
    end
    
    % Handle missing values
    if ismissing(h0_val), h0_val = ''; end
    if ismissing(h1_val), h1_val = ''; end
    if ismissing(h2_val), h2_val = ''; end
    if ismissing(h3_val), h3_val = ''; end
    
    % Build title parts
    title_parts = {};
    
    if ~isempty(h0_val) && ~strcmp(h0_val, '')
        title_parts{end+1} = h0_val;
    end
    
    if ~isempty(h2_val) && ~strcmp(h2_val, '')
        title_parts{end+1} = h2_val;
    end
    
    if ~isempty(h3_val) && ~strcmp(h3_val, '')
        title_parts{end+1} = h3_val;
    end
    
    % Create full title
    if ~isempty(title_parts)
        title_info.full_title = strjoin(title_parts, ' - ');
    else
        title_info.full_title = sprintf('Column %d', col_idx);
    end
    
    % Create ylabel
    if ~isempty(h1_val) && ~strcmp(h1_val, '')
        title_info.ylabel = sprintf('Value %s', h1_val);
    else
        title_info.ylabel = 'Value';
    end
end

function safe_name = create_safe_filename(title, col_idx)
    % Create safe filename from title
    safe_name = regexprep(title, '[^\w\s-]', '');
    safe_name = regexprep(safe_name, '\s+', '_');
    safe_name = sprintf('col%02d_%s', col_idx, safe_name);
    safe_name = safe_name(1:min(80, length(safe_name)));
end

function stats = calculate_stats(data)
    % Calculate basic statistics
    stats.mean = mean(data, 'omitnan');
    stats.median = median(data, 'omitnan');
    stats.std = std(data, 'omitnan');
    stats.min = min(data);
    stats.max = max(data);
    stats.n = sum(~isnan(data));
end
