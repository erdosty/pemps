clc
close all

% Specify the folder containing the immunofluorescence images
folder = 'positive-con';

% Get a list of all image files in the folder
fileList = dir(fullfile(folder, '*.tif'));

% Loop over each image file
for i = 1:numel(fileList)
    % Read the image
    filename = fullfile(folder, fileList(i).name);
    image = imread(filename);
    
    % Convert the image to grayscale
    grayImage = im2gray(image);
    
    % Threshold the image to create a binary mask
    threshold = graythresh(grayImage);
    binaryImage = imbinarize(grayImage, threshold);
    
    % Clean up the binary mask (optional)
    % Perform morphological operations such as erosion or dilation to remove noise or refine the mask
    cleanBinaryImage = bwareaopen(binaryImage, 500); % Remove small objects
    
    % Measure the properties of the binary mask regions
    props = regionprops(cleanBinaryImage, grayImage, 'Area', 'MeanIntensity');
    
   % Calculate the product of area and average intensity
    product = [props.Area]' .* [props.MeanIntensity]';
    
    % Calculate the weigthed pixel intensity
    weighted = ([props.Area]' .* [props.MeanIntensity]')./[props.Area]';
    
    % Display the binary mask with overlaid region properties
    overlayedImage = imoverlay(image, cleanBinaryImage, 'g'); % Green overlay
    figure;
    imshow(overlayedImage);
    title(['Masked Image ', num2str(i)]);
    
    % Display the average intensity, area, and product for each region
    for j = 1:numel(props)
        fprintf('Region %d:\n', j);
        fprintf('   Average Intensity: %.2f\n', props(j).MeanIntensity);
        fprintf('   Area: %d pixels\n', props(j).Area);
        fprintf('   Product: %.2f\n', product(j));
        fprintf('   Weighted Intensity: %.2f\n', weighted(j));
    end
    
    % Create a table with the measurements and product
    measurements = table([props.Area]', [props.MeanIntensity]', product, weighted, 'VariableNames', {'Area', 'AverageIntensity', 'TotalPixelValue','WeightedPixelValue'});
    
    % Generate a unique filename for each image
    [~, name, ~] = fileparts(filename);
    outputFilename = fullfile(folder, [name, '_measurements.xlsx']);
    
    % Save the table to an Excel file
    writetable(measurements, outputFilename);
    disp(['Measurements saved to ', outputFilename]);
end