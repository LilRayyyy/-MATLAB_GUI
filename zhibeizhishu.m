function create_vegetation_index_gui()
    % 创建一个UI图形窗口
    fig = uifigure('Name', '植被指数分析器', 'Position', [100, 100, 800, 750]);

    % 滑块上方的说明标签
    info_label = uilabel(fig, 'Position', [150, 640, 500, 20], 'Text', '若不移动滑块位置则按系统指定最佳阈值进行计算');

    % 添加选择图像按钮
    btn = uibutton(fig, 'push', 'Text', '选择图像', 'Position', [300, 680, 200, 40], 'ButtonPushedFcn', @(btn, event) select_images(btn));

    % 添加滑块
    sld = uislider(fig, 'Position', [150, 620, 500, 3], 'Limits', [0 1], 'Value', 0.65);
    sld_label = uilabel(fig, 'Position', [660, 610, 100, 20], 'Text', '阈值: 0.65');

    % 滑块是否被手动更改的标志
    slider_changed = false;

    % 分析结果
    lbl = uitextarea(fig, 'Position', [20, 20, 760, 500], 'Editable', 'off', 'FontSize', 14);

    % 选择生成对比图的复选框
    generate_comparison_checkbox = uicheckbox(fig, 'Position', [300, 570, 200, 20], 'Text', '生成对比图');
    
    % 选择保存Excel表格的复选框
    save_excel_checkbox = uicheckbox(fig, 'Position', [300, 540, 200, 20], 'Text', '保存Excel表格');

    function select_images(btn)
        % 可以选择多个图像文件
        [files, path] = uigetfile({'*.jpg;*.png;*.bmp;*.webp', '图像文件 (*.jpg, *.png, *.bmp, *.webp)'}, '选择图像', 'MultiSelect', 'on');
        
        if isequal(files, 0)
            return; % 取消选择
        end

        if ischar(files)
            files = {files}; % 如果只选择了一个文件，将其转为单元格数组
        end
        
        % 获取上一级文件夹的名称
        [~, parent_folder] = fileparts(path);

        % 初始化结果字符串
        results = '';
        % 初始化结果表格数据
        data = cell(length(files), 3);

        % 循环处理每个图像文件
        for i = 1:length(files)
            % 构建图像文件的完整路径
            imagePath = fullfile(path, files{i});
            
            try
                % 调用分析函数
                if slider_changed
                    % 使用用户手动调整的阈值
                    thresh = sld.Value;
                    [gray_avg, color_avg, gray_img, binary_mask] = process_with_threshold(imagePath, thresh);
                    best_thresh = thresh; % 将手动阈值作为最佳阈值
                else
                    % 自动寻找最佳阈值
                    [best_thresh, gray_avg, color_avg, gray_img, binary_mask] = find_best_threshold(imagePath, sld.Value);
                end

                % 更新结果字符串
                results = sprintf('%s\n%s: 最佳阈值: %.2f, 灰度图像叶子区域的平均像素值: %.2f, 原图像叶子区域的平均像素值: %.2f', results, files{i}, best_thresh, gray_avg, color_avg);
                
                % 保存结果到数据表
                data{i, 1} = files{i};
                data{i, 2} = gray_avg;
                data{i, 3} = color_avg;

                % 根据复选框状态决定是否生成对比图
                if generate_comparison_checkbox.Value
                    fig_img = uifigure('Name', sprintf('图像对比 - %s', files{i}), 'Position', [100, 100, 900, 500]);
                    ax1 = uiaxes(fig_img, 'Position', [50, 50, 350, 350]);
                    ax2 = uiaxes(fig_img, 'Position', [500, 50, 350, 350]);

                    % 显示结果图像
                    imshow(gray_img, 'Parent', ax1);
                    title(ax1, '灰度图像');
                    imshow(binary_mask, 'Parent', ax2);
                    title(ax2, '提取的叶子区域');
                end
            catch ME
                % 捕捉错误并更新结果字符串
                results = sprintf('%s\n%s: 处理失败 - %s', results, files{i}, ME.message);
            end
        end

        % 更新标签显示结果
        lbl.Value = results;

        % 输出结果到Excel表格
        if save_excel_checkbox.Value
            output_filename = fullfile(path, sprintf('%s_植被指数结果.xlsx', parent_folder));
            data_table = cell2table(data, 'VariableNames', {'图像名称', '灰度平均像素值', '原图平均像素值'});
            writetable(data_table, output_filename);
            uialert(fig, sprintf('结果已保存到 %s', output_filename), '完成');
        end

        % 重置滑块更改标志
        slider_changed = false;
    end

    function [best_thresh, gray_avg, color_avg, gray_img, binary_mask] = find_best_threshold(imagePath, initial_thresh)
        % 读取输入图像
        img = imread(imagePath);
        
        % 确保图像是RGB图像
        if size(img, 3) == 3
            % 转换为灰度图像
            gray_img = rgb2gray(img);
        elseif size(img, 3) == 1
            gray_img = img;
        else
            error('输入图像格式不支持。');
        end
        
        % 初始化参数
        best_thresh = initial_thresh;
        best_score = inf;
        thresh_values = max(0.1, initial_thresh-0.2):0.05:min(0.9, initial_thresh+0.2);

        % 尝试不同的阈值，选择最佳阈值
        for thresh = thresh_values
            T = adaptthresh(gray_img, thresh);
            binary_mask = imbinarize(gray_img, T);

            % 计算连通性（或其他指标）
            leaf_pixels = gray_img(binary_mask);
            if isempty(leaf_pixels)
                continue;
            end

            % 计算图像质量指标（这里简单使用方差作为示例）
            score = var(double(leaf_pixels));
            if score < best_score
                best_score = score;
                best_thresh = thresh;
            end
        end
        
        % 使用最佳阈值重新计算二值图像
        T = adaptthresh(gray_img, best_thresh);
        binary_mask = imbinarize(gray_img, T);
        
        % 提取灰度图像中叶子区域的像素值
        gray_leaf_pixels = gray_img(binary_mask);
        
        % 计算灰度图像叶子区域的平均像素值
        if isempty(gray_leaf_pixels)
            error('没有检测到叶子区域。请调整阈值。');
        end
        gray_avg = mean(gray_leaf_pixels);
        
        % 提取原始图像中叶子区域的像素值
        color_leaf_pixels = img(repmat(binary_mask, [1 1 3]));
        color_leaf_pixels = reshape(color_leaf_pixels, [], 3);
        color_avg = mean(color_leaf_pixels(:));
    end

    function [gray_avg, color_avg, gray_img, binary_mask] = process_with_threshold(imagePath, thresh)
        % 读取输入图像
        img = imread(imagePath);
        
        % 确保图像是RGB图像
        if size(img, 3) == 3
            % 转换为灰度图像
            gray_img = rgb2gray(img);
        elseif size(img, 3) == 1
            gray_img = img;
        else
            error('输入图像格式不支持。');
        end
        
        % 使用指定阈值来提取叶子区域
        T = adaptthresh(gray_img, thresh);
        binary_mask = imbinarize(gray_img, T);
        
        % 提取灰度图像中叶子区域的像素值
        gray_leaf_pixels = gray_img(binary_mask);
        
        % 计算灰度图像叶子区域的平均像素值
        if isempty(gray_leaf_pixels)
            error('没有检测到叶子区域。请调整阈值。');
        end
        gray_avg = mean(gray_leaf_pixels);
        
        % 提取原始图像中叶子区域的像素值
        color_leaf_pixels = img(repmat(binary_mask, [1 1 3]));
        color_leaf_pixels = reshape(color_leaf_pixels, [], 3);
        color_avg = mean(color_leaf_pixels(:));
    end

    % 更新滑块的值标签并设置滑块更改标志
    addlistener(sld, 'ValueChanged', @(sld, event) update_thresh_label_and_flag(sld, sld_label));
    
    function update_thresh_label_and_flag(sld, lbl)
        lbl.Text = sprintf('阈值: %.2f', sld.Value);
        slider_changed = true;
    end
end
