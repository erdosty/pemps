clc
close all
% Specify the folder containing the immunofluorescence images
folder = 'microglia-janus';

% Create a new Excel workbook
outputWorkbook = fullfile(folder, 'BranchAnalysisResults.xlsx');
excelApp = actxserver('Excel.Application');
workbook = excelApp.Workbooks.Add();
excelApp.Visible = 1;

% Get a list of all image files in the folder
fileList = dir(fullfile(folder, '*.tif'));

% Loop over each image file
for i = 1:numel(fileList)
    % Read the image
    filename = fullfile(folder, fileList(i).name);
    image = imread(filename);

    % Apply any preprocessing steps, such as noise reduction or contrast enhancement

    % Thresholding and Binarization
    % threshold = graythresh(image);
    % binaryImage = imbinarize(image, threshold);
    % figure
    % imshow(binaryImage);
    bin=50;
    sigma=2;
    imgfiltered = imgaussfilt(image,sigma);
    [imghist]=imhist(imgfiltered,bin);

    imgtriangle=triangle_th(imghist,bin);

    binaryImage=imbinarize(imgfiltered,imgtriangle);

    % Fill the holes

    fillimage = imfill(binaryImage,"holes");
%     figure
%     imshow(fillimage);

    openimage = bwareaopen(fillimage, 200);
%     figure
%     imshow(openimage);

    % % Skeletonization

%     remimage = bwmorph(openimage,'remove');
%     figure
%     imshow(remimage);

    skeletonImage = bwmorph(openimage,'skel',Inf);
%     figure
%     imshow(skeletonImage);

    % % Branch Analysis

    % branchPoints = bwmorph(fillImage, 'branchpoints');
    % endPoints = bwmorph(fillImage, 'endpoints');
    % branchImage = skeletonImage & ~endPoints;
    % 
    % % Counting and Visualization
    % numBranches = nnz(branchImage);
    % numEndPoints = nnz(endPoints);
    % numTotalPoints = numBranches + numEndPoints;
    % 
    % % Display the results
    % figure
    % imshow(image);
    % hold on;
    % [B, L] = bwboundaries(binaryImage, 'noholes');
    % for k = 1:length(B)
    %     boundary = B{k};
    %     plot(boundary(:,2), boundary(:,1), 'r', 'LineWidth', 2);
    % end
    % hold off;
    % 
    % fprintf('Number of branches: %d\n', numBranches);
    % fprintf('Number of end points: %d\n', numEndPoints);
    % fprintf('Total number of skeleton points: %d\n', numTotalPoints);
    
    % Identify branch points and endpoints
    branchPoints = bwmorph(skeletonImage, 'branchpoints');
    endpoints = bwmorph(skeletonImage, 'endpoints');
    
    % Label the branches and assign branch numbers to each cell
    branchImage = skeletonImage & ~endpoints;
    branchLabels = bwlabel(branchImage);
    cellLabels = branchLabels;
    cellLabels(~branchPoints) = 0;
    
    % Calculate the number of branches for each cell
    numBranches = max(cellLabels(:));
    
    % Display the skeletonized image with branches, endpoints, and cell labels
    branchOverlay = imoverlay(image, skeletonImage, 'g'); % Green overlay for branches
    endpointOverlay = imoverlay(branchOverlay, endpoints, 'r'); % Red overlay for endpoints
    cellOverlay = imoverlay(endpointOverlay, cellLabels, 'm'); % Magenta overlay for cell labels
    figure;
    imshow(cellOverlay);
    title(['Branch Analysis for Image ', num2str(i)]);

    % Display the branch number for each cell and calculate branch lengths
    fprintf('Image: %s\n', filename);
    fprintf('Number of Cells: %d\n', numBranches);
    lengthLabels = imsubtract(branchLabels, cellLabels);
    celllength = regionprops(lengthLabels, 'area');
    branchLengthspixel = (struct2array(celllength)).';
    branchLengths = (branchLengthspixel./2048).*581.25;
    for j = 1:numBranches
        cellMask = cellLabels == j;
%         branchLengths(j) = sum(skeletonImage(cellMask));
        fprintf('Cell %d: %d branches, Length: %.2f um\n', j, nnz(cellMask), branchLengths(j));
    end
    
    % Create a new worksheet in the Excel workbook
    worksheet = workbook.Sheets.Add([], workbook.Sheets.Item(workbook.Sheets.Count));
    worksheet.Name = ['Image', num2str(i)];
    
    % Write the results to the worksheet
    results = cell(numBranches+1, 3);
    results{1, 1} = 'Cell';
    results{1, 2} = 'Number of Branches';
    results{1, 3} = 'Branch Length (micrometers)';
    for j = 1:numBranches
        results{j+1, 1} = ['Cell', num2str(j)];
        results{j+1, 2} = nnz(cellLabels == j);
        results{j+1, 3} = branchLengths(j);
    end
    range = ['A1:C', num2str(numBranches+1)];
    worksheet.Range(range).Value = results;
end

% Save and close the Excel workbook
workbook.SaveAs(outputWorkbook);
excelApp.Quit();