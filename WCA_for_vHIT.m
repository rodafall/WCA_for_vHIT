% Initial Setup
clc;                     % Clear the command window.
clear;                   % Clear variables from the workspace.
clear all;               % Clear all variables and functions from memory.

% Data Importation
RAW = readtable('SubjectNumber.xlsx');  % Read the Excel file of one subject into a table named RAW.
vHIT = table2cell(RAW);            % Convert the table RAW into a cell array named vHIT.
head_vHIT_arr = vHIT(1:size(vHIT,1),83);  % Extract column 83 (head velocity data) from vHIT.
eye_vHIT_arr = vHIT(1:size(vHIT,1),84);   % Extract column 84 (eye velocity data) from vHIT.
head_vHIT_mat = [];             % Initialize an empty matrix for head velocities.
eye_vHIT_mat = [];              % Initialize an empty matrix for eye velocities.
head = {};                      % Initialize an empty cell array for processed head velocities.
eye = {};                       % Initialize an empty cell array for processed eye velocities.

% vHIT Data Extraction (Head and Eye Velocities)
for i = 1:size(head_vHIT_arr,1)    % Loop through each row of head_vHIT_arr.
    x = regexprep(head_vHIT_arr{i,1},',',' ');  % Replace commas with spaces in the string x.
    x = str2num(x);                 % Convert the string x into a numerical array.
    head_vHIT_mat = [head_vHIT_mat; x]; % Append the numerical array x to head_vHIT_mat.
end

for i = 1:size(eye_vHIT_arr,1)     % Loop through each row of eye_vHIT_arr.
    y = regexprep(eye_vHIT_arr{i,1},',',' ');   % Replace commas with spaces in the string y.
    y = str2num(y);                 % Convert the string y into a numerical array.
    eye_vHIT_mat = [eye_vHIT_mat; y]; % Append the numerical array y to eye_vHIT_mat.
end

% Time Processing
fs = 250;                          % Define the sampling frequency as 250 Hz used in ICS Impulse.
ts = 1/fs;                         % Calculate the sampling period ts.
head_vHIT = head_vHIT_mat';        % Transpose head_vHIT_mat for processing.
eye_vHIT = eye_vHIT_mat';          % Transpose eye_vHIT_mat for processing.
t_head_vHIT = ts*(1:size(head_vHIT,1)); % Create a time vector for head velocity data.
t_eye_vHIT = ts*(1:size(eye_vHIT,1));   % Create a time vector for eye velocity data.

% Wavelet Coherence Analysis
for i=1:size(eye_vHIT,2)          % Loop through each column of eye_vHIT.
    eye{i} = eye_vHIT(:,i)';      % Extract and transpose each column of eye_vHIT.
    head{i} = head_vHIT(:,i)';    % Extract and transpose each column of head_vHIT.
    eyen{i} = normalize(eye{i});  % Normalize the eye velocity data.
    headn{i} = normalize(head{i});% Normalize the head velocity data.
    t=t_eye_vHIT;                 % Set the time vector to t_eye_vHIT.
    [wcoh{i},wcs{i},f] = wcoherence(headn{i},eyen{i},250); % Compute wavelet coherence and cross-spectra.
    avg{i} = mean(wcoh{i},2);     % Calculate the mean coherence over time for each frequency.
    position{i} = find(avg{i} > 0.9); % Find the positions where coherence is greater than 0.9.
    highfp{i} = min(position{i});  % Find the first occurrence of high coherence.
    highf{i} = f(highfp{i});       % Get the corresponding frequency for high coherence.
    wcoherence(headn{i},eyen{i},250); % Plot the wavelet coherence for visualization.
end

% Exporting Results
A1 = highf';                       % Transpose the high coherence frequencies for export.
A2 = [A1 RAW(1:size(A1,1),65) RAW(1:size(A1,1),85)]; % Combine the data for export.
writetable(A2,'CoherentFrequency.xlsx');       % Write the combined data to an Excel file named 'CoherentFrequency.xlsx'.
