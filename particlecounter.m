close all;
clc;
% Opening Tiff File
Directory = 'C:\Users\erdost\Desktop\PHN Images for Analysis\deneme'; %degistir
mkdir Mask
save_dir='C:\Users\erdost\Desktop\PHN Images for Analysis\Mask'; %degistir
Folder=Directory;
files=dir(Directory);
goruntu="on";
%% Waitbar
WW=waitbar(0,'Image Processing Starts');
N=length(files);
%% Processing Loop
for i=3:N
%% Waitbar
  rtnum=num2str(i);
  RTnum=num2str(N);
  rtnum=strcat('Current Image Number: ',{' '},rtnum);
  RTnum=strcat('Total Number of Images: ',RTnum);
  disp=string({rtnum,RTnum});
  waitbar(i/N,WW,disp)
%% File acquisition
filename = files(i);
File=filename.name;
[image1,map1] = imread(fullfile(Folder, File),'TIF',1);
%% Create Mask  

image2 = im2gray(image1);
% [counts,x] = imhist(image2);
% T = otsuthresh(counts); % Otsu Threshold
imageBW = imbinarize(image2,0.5);

%% Pixel extraction and watershed functions

pixel=1000; % PIXEL^2 
imgBW2 = bwareaopen(imageBW,pixel);

D = bwdist(~imgBW2);
D = -D;
mask = imextendedmin(D,2);
D2 = imimposemin(D,mask);
L = watershed(D2);
L(~imgBW2) = 0;
imgL = imbinarize(L,0.0001);

%% Display
if goruntu=="on"
    figure(1)
    subplot(2,2,1), imshow(image2)
    subplot(2,2,2), imshow(imageBW)
    subplot(2,2,3), imshow(imgBW2)
    subplot(2,2,4), imshow(imgL)
    fig_name=strcat(File,'_panel','.tif');
    saveas(gcf,fullfile(save_dir,fig_name))
end
%% Image Save Cleared
im_name_particle=strcat(File,'_particle_count','.tif');
imwrite(imgL,fullfile(save_dir,im_name_particle),'tif')

%% Processing the Table
[labeledImage1, numBlobs1] = bwlabel(imgL);
props1 = regionprops(labeledImage1,imgL,'Area','Circularity','MajorAxisLength','MinorAxisLength');
allJanusAreas = [props1.Area];
allJanusCirc = [props1.Circularity];
allJanusMaxDia = [props1.MajorAxisLength];
allJanusMinDia = [props1.MinorAxisLength];
AvJanusDia = ((allJanusMinDia + allJanusMaxDia)/2);

header=["Area","Circularity","Max Diameter","Min Diameter","Average Diameter"];
Tablo1 = [allJanusAreas; allJanusCirc; allJanusMaxDia; allJanusMinDia; AvJanusDia].';
Tablo=[header;Tablo1];
Tablo=table(Tablo);
writetable(Tablo,strcat(File,'_janus','.xlsx'),'Sheet',1)

end
close(WW)
close(figure(1))


