#!/bin/bash

# スクリプトのいずれかのコマンドが失敗した場合に終了します。
set -e

# --- 設定項目 ---
# ご自身のSnowflake環境に合わせて以下の変数を設定してください。

# Snowflakeアカウント名 (例: xy12345.ap-northeast-1.aws)
SNOWFLAKE_ACCOUNT="<your_account_name>"

# Snowflakeユーザー名
SNOWFLAKE_USER="<your_username>"

# 使用するウェアハウス
SNOWFLAKE_WAREHOUSE="<your_warehouse>"

# 対象のデータベース
SNOWFLAKE_DATABASE="<your_database>"

# 対象のスキーマ
SNOWFLAKE_SCHEMA="<your_schema>"

# 実行したいSQLクエリ
# ヒアドキュメント(<<SQL)内にクエリを記述してください。
read -r -d '' SQL_QUERY << SQL
SELECT * FROM your_table LIMIT 100;
SQL

# 出力するCSVファイル名
OUTPUT_FILE="output.csv"

# --- スクリプト本体 ---

echo "Snowflakeへの接続を開始します..."

# snowsqlコマンドを実行
# SNOWSQL_PWD環境変数にパスワードを設定すると、パスワードプロンプトを省略できます。
# 例: export SNOWSQL_PWD='your_password'
snowsql \
  -a "${SNOWFLAKE_ACCOUNT}" \
  -u "${SNOWFLAKE_USER}" \
  -w "${SNOWFLAKE_WAREHOUSE}" \
  -d "${SNOWFLAKE_DATABASE}" \
  -s "${SNOWFLAKE_SCHEMA}" \
  -q "${SQL_QUERY}" \
  -o output_format=csv \
  -o header=true \
  -o timing=false \
  -o friendly=false \
  > "${OUTPUT_FILE}"

echo "クエリが正常に実行され、結果が'${OUTPUT_FILE}'に出力されました。"
