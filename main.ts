import { createBot, getBotIdFromToken, startBot, Intents, CreateSlashApplicationCommand, Bot, Interaction, InteractionResponseTypes } from "@discordeno/mod.ts";
import { GoogleSpreadsheet } from "npm:google-spreadsheet@4.1.4";
import { JWT } from "npm:google-auth-library@9.15.1";

import "$std/dotenv/load.ts"

interface SlashCommand {
    info: CreateSlashApplicationCommand;
    response(bot: Bot, interaction: Interaction): Promise<void>;
};

// Botのトークンを.envから取得
const BotToken: string = Deno.env.get("BOT_TOKEN")!;

const HelloCommand: SlashCommand = {
    // コマンド情報
    info: {
        name: "chaofan",
        description: "チャーハン！！！！！"
    },
    // コマンド内容
    response: async (bot, interaction) => {
        return await bot.helpers.sendInteractionResponse(interaction.id, interaction.token, {
            type: InteractionResponseTypes.ChannelMessageWithSource,
            data: {
                content: "チャーハン！！！！！！！！！！！！！！！！！！！",
                flags: 1 << 6
            }
        });
    }
}

const serviceAccountAuth = new JWT({
    email: Deno.env.get("GOOGLE_EMAIL"),
    key: Deno.env.get("GOOGLE_SECRET")!.replace(/\\n/g, "\n"),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(Deno.env.get("DOC_ID"), serviceAccountAuth);
        
const maxRow = 70

const UpdateWarPotentialCommand: SlashCommand = {
    info: {
        name: "update_war_potential",
        type: 1,
        description: "兵士戦力を更新します。",
        options: [
            {
                "name": "war_potential_k",
                "description": "兵士戦力値(k)",
                "type": 4,
                "required": true
            },
            {
                "name": "is_sub_account",
                "description": "サブアカウントの戦力を更新する場合にTrueを設定",
                "type": 5,
                "required": false
            }
        ]
    },
    response: async (bot, interaction) => {
        await doc.loadInfo();
        const sheet = doc.sheetsById[Deno.env.get("SHEET_ID")];
        
        await sheet.loadCells(`A2:E${maxRow}`);
        var isUpdateSucceeded = false;
        var lastI = 0;
        var userDendenName = '（誰だよ）';
        var beforeWarPotential = -1;
        const warPotential = interaction.data.options[0].value;
        const isSubAccount = interaction.data.options[1] && interaction.data.options[1].value;
        for (var i = 1; i < maxRow; i++) {
            const discordIdCell = sheet.getCell(i, 4);
            if (interaction.user.username == discordIdCell.value) {
                var updateTargetIndex = i;
                if (lastI > 0) { // サブアカウントがあるユーザーの更新処理
                    if (isSubAccount) {
                        updateTargetIndex = sheet.getCell(lastI, 1).value < sheet.getCell(i, 1).value ? lastI : i;
                    } else {
                        updateTargetIndex = sheet.getCell(lastI, 1).value > sheet.getCell(i, 1).value ? lastI : i;
                    }
                    isUpdateSucceeded = true;
                    userDendenName = sheet.getCell(updateTargetIndex, 0).value;
                    beforeWarPotential = await updateWarPotentialCell(warPotential, sheet, updateTargetIndex);

                    lastI = updateTargetIndex;

                    break;
                } else {
                    lastI = i;
                }
            }
        }

        if (!isUpdateSucceeded && lastI > 0) { // サブアカウントがないユーザーの更新処理
            isUpdateSucceeded = true;
            userDendenName = sheet.getCell(lastI, 0).value;
            beforeWarPotential = await updateWarPotentialCell(warPotential, sheet, lastI);
        }

        var response = '';
        if (isUpdateSucceeded) {
            response = `${userDendenName}さんの戦力を${beforeWarPotential}kから${warPotential}kに書き換えました！`;
        } else {
            response = 'エラーデス！！\nDiscordのユーザーIDと一致するスプレッドシートのデータが見つかりませんでした。\n事前にスプレッドシートにIDを登録してください。';
        }

        return await bot.helpers.sendFollowupMessage(interaction.token, {
            type: InteractionResponseTypes.ChannelMessageWithSource,
            data: {
                content: response,
                flags: 1 << 6
            }
        });
    }
}

async function handleError(e, token) {
    await bot.helpers.sendFollowupMessage(token, {
        type: InteractionResponseTypes.ChannelMessageWithSource,
        data: {
            content: 'サーバー内部でエラーが発生しました。',
            flags: 1 << 6
        }
    });
    console.error(e);
};

async function updateWarPotentialCell(warPotential, sheet, i) {
    const warPotentialCell = sheet.getCell(i, 1);
    const beforeWarPotential = warPotentialCell.value;
    warPotentialCell.value = warPotential;

    const epoch = Date.UTC(1899, 11, 30);
    const jst = 9 * 60 * 60 * 1000;
    const date = (new Date().getTime() + jst - epoch) / (1000 * 60 * 60 * 24);
    const updatedAtCell = sheet.getCell(i, 2);
    updatedAtCell.value =  date;

    await sheet.saveUpdatedCells();

    return beforeWarPotential;
};

const RegisterDiscordIdCommand: SlashCommand = {
    info: {
        name: "register_discord_id",
        type: 1,
        description: "スプレッドシートにDiscord IDを登録します。これを行うことでDiscordから戦力を更新できるようになります。",
        options: [
            {
                "name": "denden_name",
                "description": "スプレッドシート上のあなたの名前です。サブアカウント可。",
                "type": 3, // string
                "required": true
            },
            {
                "name": "should_update",
                "description": "Trueなら、既にIDが登録されていた場合に書き換えを行う。",
                "type": 5, // boolean
                "required": false
            }
        ]
    },
    response: async (bot, interaction) => {
        await doc.loadInfo();
        const sheet = doc.sheetsById[Deno.env.get("SHEET_ID")];

        await sheet.loadCells(`A2:E${maxRow}`);
        var isUpdateSucceeded = false;
        var isUpdateFailed = false;
        const givenName = interaction.data.options[0].value;
        const shouldUpdate = interaction.data.options[1] && interaction.data.options[1].value;
        const userId = interaction.user.username;
        for (var i = 1; i < maxRow; i++) {
            const dendenNameCell = sheet.getCell(i, 0);
            if (givenName == dendenNameCell.value) {
                const discordIdCell = sheet.getCell(i, 4);
                if (!shouldUpdate && discordIdCell.value && discordIdCell.value.length) {
                    isUpdateFailed = true;
                    break;
                }
                isUpdateSucceeded = true;
                discordIdCell.value = userId;
                await sheet.saveUpdatedCells();

                break;
            }
        }

        var response = '';
        if (isUpdateSucceeded) {
            response = "IDを登録しました。";
        } else if (isUpdateFailed) {
            response = "エラーデス！！\n既にIDが登録されていたため登録に失敗しました。\nIDを更新する場合にはshould_updateオプションにTrueを渡してください。"
        } else {
            response = "エラーデス！！\n該当する名前がスプレッドシート上に存在しませんでした。名前の表記揺れなどがないかご確認ください。"
        }

        return await bot.helpers.sendFollowupMessage(interaction.token, {
            type: InteractionResponseTypes.ChannelMessageWithSource,
            data: {
                content: response,
                flags: 1 << 6
            }
        });
    }
}

const RankCommand: SlashCommand = {
    info: {
        name: "rank",
        type: 1,
        description: "あなたの兵士戦力順位と部隊名を教えます。登録されている場合にはサブアカウントについても回答します。",
    },
    response: async (bot, interaction) => {
        await doc.loadInfo();
        const inputSheet = doc.sheetsById[Deno.env.get("SHEET_ID")];    
        const rankingSheet = doc.sheetsById[Deno.env.get("RANKING_SHEET_ID")];
        
        await inputSheet.loadCells(`A2:E${maxRow}`);
        var dendenNames: string[] = [];
        for (var i = 1; i < maxRow; i++) {
            const discordIdCell = inputSheet.getCell(i, 4);
            if (interaction.user.username == discordIdCell.value) {
                dendenNames.push(inputSheet.getCell(i, 0).value);
            }
        }
        if (!dendenNames.length) {
            return await bot.helpers.sendFollowupMessage(interaction.token, {
                type: InteractionResponseTypes.ChannelMessageWithSource,
                data: {
                    content: 'エラーデス！！\nDiscordのユーザーIDと一致するスプレッドシートのデータが見つかりませんでした。\n事前にスプレッドシートにIDを登録してください。',
                    flags: 1 << 6
                }
            });
        }

        await rankingSheet.loadCells(`B2:D${maxRow}`);
        var results: string[] = [];
        for (var i = 1; i < maxRow; i++) {
            const dendenNameCell = rankingSheet.getCell(i, 2);
            for (const dendenName of dendenNames) {
                if (dendenName == dendenNameCell.value) {
                    const rank = rankingSheet.getCell(i, 1).value;
                    var unit = '労働';
                    if (rank <= 25) {
                        unit = '決死';
                    } else if (rank <= 50) {
                        unit = '調査';
                    }
                    results.push(`${dendenName}さんは${unit}${rank}`);
                }
            }
        }

        var response = '';
        if (results.length) {
            response = `${results.join('、')}です。`;
        } else {
            response = 'エラーデス！！\nDiscordのユーザーIDと一致するスプレッドシートのデータが見つかりませんでした。\n事前にスプレッドシートにIDを登録してください。';
        }
        return await bot.helpers.sendFollowupMessage(interaction.token, {
            type: InteractionResponseTypes.ChannelMessageWithSource,
            data: {
                content: response,
                flags: 1 << 6
            }
        });
    }
}

// ボットの作成
const bot = createBot({
    token: BotToken,
    botId: getBotIdFromToken(BotToken) as bigint,

    intents: Intents.Guilds | Intents.GuildMessages,

    // イベント発火時に実行する関数など
    events: {
        // 起動時
        ready: (_bot, payload) => {
            console.log(`${payload.user.username} is ready!`);
        },
        interactionCreate:async  (_bot, interaction) => {
            console.log(`/${interaction.data.name} performed by ${interaction.user.username}.`, interaction.data);
            if (interaction.data.name == "chaofan") {
                await HelloCommand.response(bot, interaction);
            } else {
                await bot.helpers.sendInteractionResponse(interaction.id, interaction.token, {
                    type: InteractionResponseTypes.DeferredChannelMessageWithSource,
                    data: {
                        flags: 1 << 6
                    }
                });
                if (interaction.data.name == "update_war_potential") {
                    await UpdateWarPotentialCommand.response(bot, interaction).catch(async (e) => { await handleError(e, interaction.token) });
                } else if (interaction.data.name == "register_discord_id") {
                    await RegisterDiscordIdCommand.response(bot, interaction).catch(async (e) => { await handleError(e, interaction.token) });
                } else if (interaction.data.name == "rank") {
                    await RankCommand.response(bot, interaction).catch(async (e) => { await handleError(e, interaction.token) });
                }
            }
        }
    }
});

// コマンドの作成
bot.helpers.createGlobalApplicationCommand(HelloCommand.info);
bot.helpers.createGlobalApplicationCommand(UpdateWarPotentialCommand.info);
bot.helpers.createGlobalApplicationCommand(RegisterDiscordIdCommand.info);
bot.helpers.createGlobalApplicationCommand(RankCommand.info);

// コマンドの登録
bot.helpers.upsertGlobalApplicationCommands([UpdateWarPotentialCommand.info, HelloCommand.info, RegisterDiscordIdCommand.info, RankCommand.info]);

await startBot(bot);
