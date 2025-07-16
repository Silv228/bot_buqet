import TelegramBot from 'node-telegram-bot-api';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

const TOKEN = '7839533817:AAG-Uo_7LETNmWFqZviDAaw9CnUAcOqphRY'

const serviceAccountAuth = new JWT({
    email: 'admin-888@botbuqets.iam.gserviceaccount.com',
    key: '-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC0jqlDhGKHmvbR\nXQCTh7qr6YzbUzXG/A/sfl7ArDPnl4GxW8XfPY5fEo8/9zegvF7fQQTFQ/7U7bVu\nW2j7LFkVBbk5kX1I+eoew9IV8x4AVf0XgR2P84aC1/RlFxKSBqkJQzcF2WByzokG\n7d3g5hWHwtlqoEOPiuI6eiEJZkmlc7qjJ4UWgK+QUOvNwafjNCWwPBze/TvYSM4I\npVAlZggAzgmyHIFN3byNOrAsLgHncb/Ve8kapu6SeBA2uAY+p9rKERF/pYwOVLYL\n/8bjNP97MsXFYCrUVgI+59VFtoTvxhUC8NaHqBv9JWLwL/rIaomFKWi4auOte6vv\nJSvh//0tAgMBAAECggEAAgHBEU0wrOIJV2hiy2WQGJjIGHGxCcGWCgJ0/t5h+nNt\nXSzeGbvT2SL1h8pWE4UkeWFZe4Kt0pF6YqCgnajvh2XcLHgPpyRsir9IchJL18jT\nvy79fY5BPA7I/wzvB3p6C1wrLkB2a4uutMoCnS92EqinV2rE4y8fkgVkoXevuid2\nvGvzXLVZ5tlP7zeVUv00/4B/dsvaRsabc0DhTeND7/qBVoRffrohLfFBWvKrzXas\nI7NlFjLEhia2Hhs3cmjO6V8+0oxRKe+qMFBeMGFtobe4Gp+95A70OsbRveqxzj9E\n/yDrdrQ9Z52qID0tgLsXWTEndlCF44K3KSst8q3XjQKBgQD6+O3srmyQ0itXfKCf\nhqdKMoupsMaA7+vj+EtO7PzBc8Pm0KGYDOyznw+73tvxE9hK0wMvPPNPzad8pvvO\nMeM1WQRFtLTIacvgV+RKJSnNL5UGfwGEJsdhmTL1vqcVtGIe5Pch9tnxNO/nf53j\nXFgHg4E66VVP/Mfakrheyant5wKBgQC4LJ6V0LVEej5Ofsj7aHaSCqTQaPGfBfmg\nwxGpsU7HMXWEElTmqHKMfqiD3qhOiFMOqMQinosmwUSO2yCjHrPQmwRkvgnrENa1\n4WX3e9sx/jnZkAacGsJjl1SYx59SYlLn+D/8fgg3lccesW4vQHxQZ2X6e/+qeM/M\nLiqJ3mwRywKBgQCiMZIB7c+34DumdKKRtkITD4t3BQmkdmlqkSKKRVor45btalOk\nomWux9MxRRu7N2oHIUvjkW5lWrEtO/VsEo2WAotiSSC0jLr3p5Wf3VighGm5Iwdl\n0nH0Pz/R1X7B5iurb6nPR2seGWoZoD33m8xAPtqbqgQ6h1DZjwycJZQubQKBgCL2\nTXCROyfxsMxD4zFepkuY+6qYkW2nu7iZ70twXk0QBYf51uYmigBDtwe5h+fIl0PM\nI9eSk0XbIIGh9XMhy+7Izq+1J7rY6nmCfVHa0ESQRzkWzzppFgfD3YpXMtZ31dc1\nWCg9YJ/0reUUt57+tdqplkFTsrgQ0RmoleiwMYG7AoGAN2n6WGzHqQRyz8LeO2x6\nH31jfzJcrHhs9qRA/hF9AdFzms/2D/WDCoW6pniUsG37Ak/TRfQHmEz7uyNy9vLH\nufLNsW3LqifqXCAG/R7C208ZmqaDk96PIPN458QzX3nC/OEX13dC2JoLY3V7aNiV\nhKH/OfyKBvea1ul+J6bM9+w=\n-----END PRIVATE KEY-----\n'.split(String.raw`\n`).join('\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});


const doc = new GoogleSpreadsheet('1fBmVYLbrtKRVmCeItyadY7VqfV1NhwpBduiYrjuDo-o', serviceAccountAuth);
await doc.loadInfo();
console.log(doc.title);
const ids = doc.sheetsByTitle['id']
const months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']

const bot = new TelegramBot(TOKEN, {
    polling: true
});

bot.on('text', async msg => {
    try {
        const month_cur = new Date().getMonth()
        const month = doc.sheetsByIndex[month_cur + 1]
        let users = []
        const rowsId = await ids.getRows()

        for (let i = 0; i < rowsId.length; i++) {
            users.push([rowsId[i].get('id'), rowsId[i].get('name'), rowsId[i].get('role')]);
        }
        console.log(msg.from.id, users)
        if (!users.find((el) => +el[0] === msg.from.id)) {
            await bot.sendMessage(msg.chat.id, 'У вас нет доступа!')
        }
        else {
            switch (msg.text) {
                case '/start':
                    let users = []
                    const rowsId = await ids.getRows()
                    for (let i = 0; i < rowsId.length; i++) {
                        users.push([rowsId[i].get('id'), rowsId[i].get('name'), rowsId[i].get('role')]);
                    }
                    if (!users.find((el) => +el[0] === msg.from.id)) {
                        await bot.sendMessage(msg.chat.id, 'У вас нет доступа!')
                    }
                    else {
                        await bot.sendMessage(msg.chat.id, 'Чтобы начать смену нажмите кнопку "Начать"', {
                            reply_markup: {
                                keyboard: [
                                    ['Начать'],
                                    ['Доход']
                                ]
                            }
                        })
                    }
                    break
                case 'Начать':
                    let users1 = []
                    const rowsId1 = await ids.getRows();
                    const monthRow = await month.getRows();
                    const col = month.headerValues.findIndex(el => el == `${new Date().toLocaleString("ru", { day: "2-digit" })}.${new Date().toLocaleString("ru", { month: "2-digit" })}`)
                    for (let i = 0; i < rowsId1.length; i++) {
                        users1.push([rowsId1[i].get('id'), rowsId1[i].get('name'), rowsId1[i].get('role')]);
                    }
                    const name = users1.find((el) => +el[0] === msg.from.id)[1]
                    const role = users1.find((el) => +el[0] === msg.from.id)[2]
                    await month.loadCells({ startRowIndex: 0, startColumnIndex: 0 })
                    const nameCol = month.headerValues.findIndex(el => el === 'имя')
                    monthRow.length = monthRow.length > 0 ? monthRow.length : 2
                    for (let i = 0; i <= monthRow.length; i++) {
                        if (month.getCell(i, nameCol).value === name) {
                            console.log(monthRow.length)
                            for (let j = 0; j < 35; j++) {
                                if (month.getCell(i + 1, col).value === null) {
                                    let cellStart = month.getCell(i, col)
                                    cellStart.value = new Date().toLocaleTimeString("ru", { timeZone: '+05:00', hour: '2-digit', minute: '2-digit' });
                                    await month.saveUpdatedCells();
                                    break
                                }
                            }
                            break
                        }
                        else {
                            if (i + 1 === monthRow.length) {
                                console.log(monthRow.length)
                                if (monthRow.length > 2) {
                                    for (let k = 0; k < month.headerValues.length; k++) {
                                        const cell = month.getCell(monthRow.length + 2, k)
                                        cell.backgroundColor = {
                                            red: 10,
                                            blue: 1,
                                            green: 10,
                                            alpha: 10
                                        }
                                    }
                                    await month.saveUpdatedCells();
                                    month.getCell(monthRow.length + 2, 0).value = name
                                    month.getCell(monthRow.length + 3, 0).value = role
                                    let cellStart = month.getCell(monthRow.length + 2, col)
                                    console.log(i, 'i')
                                    cellStart.value = new Date().toLocaleTimeString("ru", { timeZone: '+05:00', hour: '2-digit', minute: '2-digit' });
                                    // cellStart.value = 'start'
                                    await month.saveUpdatedCells();
                                }
                                else {
                                    monthRow.length = 2
                                    for (let k = 0; k < month.headerValues.length; k++) {
                                        const cell = month.getCell(monthRow.length, k)
                                        cell.backgroundColor = {
                                            red: 40,
                                            blue: 1,
                                            green: 10,
                                            alpha: 10
                                        }
                                    }
                                    await month.saveUpdatedCells();
                                    month.getCell(monthRow.length, 0).value = name
                                    month.getCell(monthRow.length + 1, 0).value = role
                                    let cellStart = month.getCell(monthRow.length, col)
                                    console.log(i, monthRow.length)
                                    cellStart.value = new Date().toLocaleTimeString("ru", { timeZone: '+05:00', hour: '2-digit', minute: '2-digit' });
                                    await month.saveUpdatedCells();
                                }
                                // await monthRow[i].save();
                                await month.saveUpdatedCells();
                            }
                            else {
                                continue
                            }
                        }

                    }

                    await bot.sendMessage(msg.chat.id, 'Чтобы закончить смену нажмите кнопку "Закончить"', {
                        reply_markup: {
                            keyboard: [
                                ['Закончить'],
                                ['Доход']
                            ]
                        }
                    })
                    break
                case 'Закончить':
                    let users2 = []
                    const rowsId2 = await ids.getRows();
                    const monthRow2 = await month.getRows();
                    const col1 = month.headerValues.findIndex(el => el == `${new Date().toLocaleString("ru", { day: "2-digit" })}.${new Date().toLocaleString("ru", { month: "2-digit" })}`)
                    for (let i = 0; i < rowsId2.length; i++) {
                        users2.push([rowsId2[i].get('id'), rowsId2[i].get('name'), rowsId2[i].get('role')]);
                    }
                    const nameCol1 = month.headerValues.findIndex(el => el === 'имя')
                    await month.loadCells({ startRowIndex: 0, startColumnIndex: 0 })
                    const summaryCol = month.headerValues.findIndex(el => el === 'сумма часов')
                    for (let i = 0; i <= monthRow2.length; i++) {
                        if (month.getCell(i, nameCol1).value === users2.find((el) => +el[0] === msg.from.id)[1]) {
                            for (let j = 1; j < 35; j++) {
                                if (month.getCell(i + 1, col1).value === null) {
                                    let cellEnd = month.getCell(i + 1, col1)
                                    cellEnd.value = new Date().toLocaleTimeString("ru", { timeZone: '+05:00', hour: '2-digit', minute: '2-digit' });
                                    await month.saveUpdatedCells();
                                    let cellDiff = month.getCell(i + 2, col1)
                                    let minutes = (month.getCell(i + 1, col1).value.split(':')[0] * 60 + month.getCell(i + 1, col1).value.split(':')[1]) - (month.getCell(i, col1).value.split(':')[0] * 60 + month.getCell(i, col1).value.split(':')[1])
                                    cellDiff.value = `${Math.floor(minutes / 60) % 99}:${minutes % 60}`
                                    console.log(`${Math.floor(minutes / 60)}:${minutes % 60}`)
                                    await month.saveUpdatedCells();
                                    const sumCell = Number(String(month.getCell(i, summaryCol).value).replace(',', '.'))
                                    let sum = 0
                                    console.log(sumCell, summaryCol)
                                    if (!sumCell) {
                                        sum = ((minutes / 60) % 99)
                                    }
                                    else {
                                        sum = sumCell + ((minutes / 60) % 99)
                                    }
                                    console.log(sum)
                                    const priceCol = month.headerValues.findIndex(el => el === 'ставка')
                                    const planCol = month.headerValues.findIndex(el => el === 'план')
                                    const factCol = month.headerValues.findIndex(el => el === 'факт')
                                    const summary_hCol = month.headerValues.findIndex(el => el === 'доход от часов')
                                    const summary_sCol = month.headerValues.findIndex(el => el === 'доход от плана')
                                    const resultCol = month.headerValues.findIndex(el => el === 'итог')
                                    let price = month.getCell(i, priceCol).value
                                    // console.log(monthRow2[i - 1], month.getCell(i, priceCol).value, month.getCell(i - 1, priceCol).value)
                                    let k1 = Number(monthRow2[i - 1].get('k1').replace(',', '.'))
                                    let k2 = Number(monthRow2[i - 1].get('k2').replace(',', '.'))
                                    let plan = month.getCell(i, planCol).value ? +month.getCell(i, planCol).value : 0
                                    let fact = +month.getCell(i, factCol).value
                                    let summary_s = !month.getCell(i, summary_sCol).value ? 0 : Number(String(month.getCell(i, summary_sCol).value).replace(',', '.'))
                                    console.log(month.getCell(i, summary_sCol).value, 'план')
                                    if (fact >= plan) {
                                        summary_s += fact * (k2 / 100)
                                    }
                                    else {
                                        summary_s += fact * (k1 / 100)
                                    }
                                    console.log(price, k1, k2, sum, summary_s)
                                    month.getCell(i, summaryCol).value = sum.toFixed(2)
                                    await month.saveUpdatedCells();
                                    month.getCell(i, summary_hCol).value = (Number(price) * sum).toFixed(2)
                                    await month.saveUpdatedCells();
                                    month.getCell(i, summary_sCol).value = (summary_s).toFixed(2)
                                    await month.saveUpdatedCells();
                                    month.getCell(i, resultCol).value = (summary_s + Number(price) * (Number(String(sum).replace(',', '.')))).toFixed(2)
                                    await month.saveUpdatedCells();
                                    break
                                }
                            }
                        }
                        else {
                        }
                    }

                    await bot.sendMessage(msg.chat.id, 'Чтобы начать смену нажмите кнопку "Начать"', {
                        reply_markup: {
                            keyboard: [
                                ['Начать'],
                                ['Доход']
                            ]
                        }
                    })
                    break
                case 'Доход':
                    let users3 = []
                    const rowsId3 = await ids.getRows();
                    await month.loadCells({ startRowIndex: 0, startColumnIndex: 0 })
                    const monthRow3 = await month.getRows();
                    const resultCol = month.headerValues.findIndex(el => el === 'итог')
                    const nameCol2 = month.headerValues.findIndex(el => el === 'имя')
                    for (let i = 0; i < rowsId3.length; i++) {
                        users3.push([rowsId3[i].get('id'), rowsId3[i].get('name')]);
                    }
                    for (let i = 0; i <= monthRow3.length; i++) {
                        if (month.getCell(i, nameCol2).value === users3.find((el) => +el[0] === msg.from.id)[1]) {
                            const sal_id = await bot.sendMessage(msg.chat.id, `Ваш доход - ${month.getCell(i, resultCol).value} рублей`)
                            console.log(sal_id.message_id)
                            setTimeout(() => {
                                bot.deleteMessage(msg.chat.id, sal_id.message_id)
                            }, 5000 * 60)
                        }
                    }

                default:
                    console.log(msg.text)
                    break
            }
        }

    } catch (error) {
        console.log(error)
    }
})

//тест перехода вермени