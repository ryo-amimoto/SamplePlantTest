% RCP実機テストの自動化スクリプト         %
% 作成者:網本　亮  作成日:2020/10/13     %

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 制御モデルのテストデータ（matファイル）をエクセルへ変換　%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
load("InputVal.mat");
load("OutputVal.mat");
InputHead = num2cell(importdata("InputIF_In_Header.xlsx"));
OutputHead = num2cell(importdata("InputIF_Out_Header.xlsx"));
InputData = vertcat(InputHead,InputVal);
OutputData = vertcat(OutputHead,OutputVal);

% モデルの入出力データを”テストファイル＿マクロ機能付き”へ出力
xlswrite("TestGraphOutput.xlsm",InputVal,"InputIF_In",'A2:J412');
xlswrite("TestGraphOutput.xlsm",OutputVal,"InputIF_Out",'A2:I412');

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% テストデータのグラフ出力をエクセルのマクロ機能で実行　   %
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% % オブジェクトの生成 
% ExcelApp = actxserver('Excel.Application'); 
% % Excel画面の表示 
% ExcelApp.Visible = 1; 
% % カレントフォルダにあるファイルを開く
% ExcelApp.Workbooks.Open(fullfile(pwd,'\TestGraphOutput.xlsm')); 
% 
% % マクロ：MakingGraphの実行。マクロは"Module1"モジュールに定義され、引数は不要です。 
% retVal = ExcelApp.Run('MakingGraph'); 
% 
% % Excelを終了し、オブジェクトを消去します。
% ExcelApp.Quit; 
% ExcelApp.release;
% 
% winopen("TestGraphOutput.xlsm");



