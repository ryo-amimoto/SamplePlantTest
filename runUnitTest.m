% RCP単体テストの自動化スクリプト         %
% 作成者:網本　亮  作成日:2020/10/6      %

simOut = sim("Mdl_Sample");     % モデルのシミュレーション実行

% モデルのテスト結果ファイルを作成
OutputData = num2cell(simOut.ScopeData(1:end,2:end));   % モデルの出力データを作成
xlswrite("TestData.xlsx",OutputData,"テスト結果",'D3:E126');    % モデルの出力データをテスト結果ファイルへ出力
% winopen("TestData.xlsx");



