import discord
import openpyxl

# Discord API認証トークンを設定する
TOKEN = '***'

intents = discord.Intents.default()  # デフォルトのIntentsオブジェクトを生成
intents.typing = True  # typingを受け取らないように

# Discordクライアントを初期化する
client = discord.Client(intents=intents)

# Discordに接続する
@client.event
async def on_ready():
    print('Logged in as {0.user}'.format(client))

    # サーバーのメンバーリストを取得する
    server = client.get_guild(***) # サーバーIDを指定する

    if server is None:
        print("サーバーを取得できませんでした。正しいサーバーIDを指定していますか？")
    else:
        members = server.members
        # メンバー情報を処理するコードを追加する

    if server is not None:
        members = server.members
        # メンバー情報を処理するコードを追加する
    else:
        print("サーバーを取得できませんでした。")

    
    members = server.members

    # Excelファイルを作成し、メンバーリストを出力する
    wb = openpyxl.Workbook()
    ws = wb.active

    # ヘッダーを書き込む
    ws.cell(row=1, column=1, value='Username')
    ws.cell(row=1, column=2, value='Discriminator')
    ws.cell(row=1, column=3, value='Roles')

    # メンバーリストを書き込む
    for i, member in enumerate(members):
        roles = [role.name for role in member.roles if role.name != "@everyone"]
        role_str = ", ".join(roles)
        ws.cell(row=i+2, column=1, value=member.name)
        ws.cell(row=i+2, column=2, value=member.discriminator)
        ws.cell(row=i+2, column=3, value=role_str)
        print(member.name)

    # Excelファイルを保存する
    wb.save('members.xlsx')

    print("Finish")
    
# Discordに接続する
client.run(TOKEN)
