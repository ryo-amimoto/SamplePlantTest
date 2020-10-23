% RCP実機テストの自動化スクリプト       %
% 作成者:網本　亮  作成日:2020/10/13    %

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 制御モデルのテストデータ（matファイル）をエクセルへ変換　%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
load("InputVal.mat");
load("OutputVal.mat");
InputHead = num2cell(importdata("InputIF_In_Header.xlsx"));
OutputHead = num2cell(importdata("InputIF_Out_Header.xlsx"));
InputData = vertcat(InputHead,InputVal);
OutputData = vertcat(OutputHead,OutputVal);

% モデルの入出力データをグラフ変換
[n,p] = size(InputVal);
[o,q] = size(OutputVal);
figure;
xi = cell2table(InputVal(1:end , 1));
yi = cell2table(InputVal(1:end , 2));
subplot(p-1,2,1),plot(xi.Var1,yi.Var1);
for s = 1:p-2
    % 入力
    xi = cell2table(InputVal(1:end , 1));
    yi = cell2table(InputVal(1:end , s+2));
    subplot(p-1,2,2*s+1),plot(xi.Var1,yi.Var1);
    % 出力
    xo = cell2table(OutputVal(1:end , 1));
    yo = cell2table(OutputVal(1:end , s+1));
    subplot(p-1,2,2*s+2),plot(xo.Var1,yo.Var1);
end

savefig("TestGraph");

% 

% % モデルの入出力データを”テストファイル＿マクロ機能付き”へ出力
% xlswrite("TestGraphOutput.xlsm",InputVal,"InputIF_In",'A2:J412');
% xlswrite("TestGraphOutput.xlsm",OutputVal,"InputIF_Out",'A2:I412');

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



