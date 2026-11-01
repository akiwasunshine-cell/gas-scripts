# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクトの状況

Google Apps Script（`.gs`）のスクリプトを管理するリポジトリです。ビルド・テスト・Lintのコマンドや、Apps Scriptへのデプロイ手段は未整備のため、決まり次第ここに追記してください。

- `summary_dashboard.gs`：スプレッドシートにバインドして使う想定。「売上データ」シートを月ごとに集計し、「月次サマリー」シートに結果と棒グラフを出力する（`runDashboard`）。`setDailyTrigger`は毎朝9時台の実行トリガーを登録する
- `email_classifier.gs`：スプレッドシートにバインドして使う想定。Gmailの「要処理」ラベルの未読メールをClaude APIで分類し、「メールログ」シートへの記録とSlack通知を行う（`runEmailClassifier`）。`setEmailClassifierTrigger`は5分おきの実行トリガーを登録する。スクリプトプロパティ`CLAUDE_API_KEY`と`SLACK_WEBHOOK_URL`が必要

`.gs`ファイルは同じApps Scriptプロジェクトに置くとグローバルスコープを共有するため、ファイル間で定数名や関数名を重複させないでください（`const`の重複は構文エラーになります）。

## Git運用ルール

コードを変更するたびに、変更をコミットしてGitHubにプッシュしてください。ユーザーから都度の指示がなくても実施します。

- 変更作業が1つ完了するごとに、コミットとプッシュをセットで行う（現在のブランチへプッシュする）
- リポジトリが未初期化、またはリモートが未設定でプッシュできない場合は、リポジトリやリモートを勝手に作成せず、その旨をユーザーに伝えて指示を仰ぐ
