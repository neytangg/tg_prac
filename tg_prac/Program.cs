using System;
using System.ComponentModel.DataAnnotations;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Extensions;
using Microsoft.Office.Interop.Excel;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using System.Data.SQLite;
using Telegram.Bot.Polling;
using System.Text;
using System.Data.Common;

namespace tg_bot
{
    class Program
    {
        static TelegramBotClient botClient = new TelegramBotClient("5703627547:AAFpUHVuN_g4M4c4XTN65lP4bLH4dFWfYYE");
        static void Main(string[] args)
        {
            Console.WriteLine($"Бот {botClient.GetMeAsync().Result.FirstName} запустився.");
            Console.OutputEncoding = UTF8Encoding.UTF8;

            using var cts = new CancellationTokenSource();
            var cancellationToken = cts.Token;
            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = { }
            };

            botClient.StartReceiving(
             UpdateHandler,
             HandleErrorAsync,
             receiverOptions,
             cancellationToken);



            Console.ReadLine();

            cts.Cancel();

            async Task HandleUpdatesAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
            {
                //
            }

            async Task UpdateHandler(ITelegramBotClient botClient, Update update, CancellationToken arg3)
            {
                var message = update.Message;
                if (update.Type == UpdateType.Message && update?.Message?.Text != null)
                {
                    if (update.Message.Type == MessageType.Text)
                    {
                        var text = update.Message.Text;
                        var id = update.Message.Chat.Id;
                        string? username = update.Message.Chat.Username;
                        var data = update.Message.Date;

                        Console.WriteLine($"{username} | {id} | {text} | {data}");


                        if (message.Text == "/start")
                        {
                            Register(message.Chat.Id.ToString(), message.Chat.Username.ToString(), DateTime.Now.ToString());
                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
                     new KeyboardButton[] {"Інженерія ПЗ", "Автоматизація"},
                     new KeyboardButton[] {"Менеджмент", "Маркетинг"},
                     new KeyboardButton[] {"Фінанси", "Облік і оподаткування"},
                     new KeyboardButton[] {"Технології ЛП", "Машинобудування"},
                     new KeyboardButton[] {"Деревообробні та МТ"}
        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть спеціальність:", replyMarkup: keyboard);
                            return;
                        }
//Інженерія ПЗ
                        if (message.Text == "Інженерія ПЗ")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"П-11", "П-12" },
            new KeyboardButton[] {"П-21", "П-22"},
            new KeyboardButton[] {"П-31", "П-32"},
            new KeyboardButton[] {"П-41", "П-42"}

        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "П-11")
                        {
                            ReadExcel(botClient, update, "П", "11");
                        }
                        if (message.Text == "П-12")
                        {
                            ReadExcel(botClient, update, "П", "12");
                        }
                        if (message.Text == "П-21")
                        {
                            ReadExcel(botClient, update, "П", "21");
                        }
                        if (message.Text == "П-22")
                        {
                            ReadExcel(botClient, update, "П", "22");
                        }
                        if (message.Text == "П-23")
                        {
                            ReadExcel(botClient, update, "П", "23");
                        }
                        if (message.Text == "П-31")
                        {
                            ReadExcel(botClient, update, "П", "31");
                        }
                        if (message.Text == "П-32")
                        {
                            ReadExcel(botClient, update, "П", "32");
                        }
                        if (message.Text == "П-41")
                        {
                            ReadExcel(botClient, update, "П", "41");
                        }
                        if (message.Text == "П-42")
                        {
                            ReadExcel(botClient, update, "П", "42");
                        }


 //Фінанси
                        if (message.Text == "Фінанси")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"Ф-11", "Ф-12"},
            new KeyboardButton[] {"Ф-21", "Ф-22"},
            new KeyboardButton[] {"Ф-31", "Ф-32"}
        })

                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "Ф-11")
                        {
                            ReadExcel(botClient, update, "Ф", "11");
                        }
                        if (message.Text == "Ф-12")
                        {
                            ReadExcel(botClient, update, "Ф", "12");
                        }
                        if (message.Text == "Ф-21")
                        {
                            ReadExcel(botClient, update, "Ф", "21");
                        }
                        if (message.Text == "Ф-22")
                        {
                            ReadExcel(botClient, update, "Ф", "22");
                        }
                        if (message.Text == "Ф-23")
                        {
                            ReadExcel(botClient, update, "Ф", "23");
                        }
                        if (message.Text == "Ф-31")
                        {
                            ReadExcel(botClient, update, "Ф", "31");
                        }
                        if (message.Text == "Ф-31")
                        {
                            ReadExcel(botClient, update, "Ф", "32");
                        }

//Автоматизація
                        if (message.Text == "Автоматизація")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"А-11", "А-12" },
            new KeyboardButton[] {"А-21", "А-22"},
            new KeyboardButton[] {"А-31", "А-32"},
            new KeyboardButton[] {"А-41", "А-42"}
        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "А-11")
                        {
                            ReadExcel(botClient, update, "А", "11");
                        }
                        if (message.Text == "А-12")
                        {
                            ReadExcel(botClient, update, "А", "12");
                        }
                        if (message.Text == "А-21")
                        {
                            ReadExcel(botClient, update, "А", "21");
                        }
                        if (message.Text == "А-22")
                        {
                            ReadExcel(botClient, update, "А", "22");
                        }
                        if (message.Text == "А-23")
                        {
                            ReadExcel(botClient, update, "А", "23");
                        }
                        if (message.Text == "А-31")
                        {
                            ReadExcel(botClient, update, "А", "31");
                        }
                        if (message.Text == "А-32")
                        {
                            ReadExcel(botClient, update, "А", "32");
                        }
                        if (message.Text == "А-41")
                        {
                            ReadExcel(botClient, update, "А", "41");
                        }
                        if (message.Text == "А-42")
                        {
                            ReadExcel(botClient, update, "А", "42");
                        }

//Менеджмент
                        if (message.Text == "Менеджмент")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"Мд-11", "Мд-12" , "Мд-31"},

        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "Мд-11")
                        {
                            ReadExcel(botClient, update, "Мд", "11");
                        }
                        if (message.Text == "Мд-12")
                        {
                            ReadExcel(botClient, update, "Мд", "12");
                        }
                        if (message.Text == "Мд-31")
                        {
                            ReadExcel(botClient, update, "Мд", "31");
                        }

//Маркетинг
                        if (message.Text == "Маркетинг")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"Е-11", "Е-21" },
            new KeyboardButton[] {"Е-31", "Е-41"},

        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "Е-11")
                        {
                            ReadExcel(botClient, update, "Е", "11");
                        }
                        if (message.Text == "Е-21")
                        {
                            ReadExcel(botClient, update, "Е", "21");
                        }
                        if (message.Text == "Е-31")
                        {
                            ReadExcel(botClient, update, "Е", "31");
                        }
                        if(message.Text == "Е-41")
                        {
                            ReadExcel(botClient, update, "Е", "41");
                        }

//Облік і оподаткування
                        if (message.Text == "Облік і оподаткування")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"Б-11", "Б-21" , "Б-31"},


        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "Б-11")
                        {
                            ReadExcel(botClient, update, "Б", "11");
                        }
                        if (message.Text == "Б-12")
                        {
                            ReadExcel(botClient, update, "Б", "12");
                        }
                        if (message.Text == "Б-31")
                        {
                            ReadExcel(botClient, update, "Б", "31");
                        }

//Технології ЛП
                        if (message.Text == "Технології ЛП")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"Мк-11", "Мк-21" , "Мк-31" , "Мк-41"},


        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "Мк-11")
                        {
                            ReadExcel(botClient, update, "Мк", "11");
                        }
                        if (message.Text == "Мк-12")
                        {
                            ReadExcel(botClient, update, "Мк", "12");
                        }
                        if (message.Text == "Мк-31")
                        {
                            ReadExcel(botClient, update, "Мк", "31");
                        }
                        if (message.Text == "Мк-41")
                        {
                            ReadExcel(botClient, update, "Мк", "41");
                        }

//Машинобудування
                        if (message.Text == "Машинобудування")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"М-11", "М-21" , "М-31" , "М-41"},


        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "М-11")
                        {
                            ReadExcel(botClient, update, "М", "11");
                        }
                        if (message.Text == "М-12")
                        {
                            ReadExcel(botClient, update, "М", "12");
                        }
                        if (message.Text == "М-31")
                        {
                            ReadExcel(botClient, update, "М", "31");
                        }
                        if (message.Text == "М-41")
                        {
                            ReadExcel(botClient, update, "М", "41");
                        }

//Деревообробні та МТ
                        if (message.Text == "Деревообробні та МТ")
                        {

                            ReplyKeyboardMarkup keyboard = new(new[]
                            {
            new KeyboardButton[] {"T-11", "T-21" , "T-31"},


        })
                            {
                                ResizeKeyboard = true
                            };
                            await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть курс:", replyMarkup: keyboard);
                            return;
                        }
                        if (message.Text == "Т-11")
                        {
                            ReadExcel(botClient, update, "Т", "11");
                        }
                        if (message.Text == "Т-12")
                        {
                            ReadExcel(botClient, update, "Т", "12");
                        }
                        if (message.Text == "Т-31")
                        {
                            ReadExcel(botClient, update, "Т", "31");
                        }

                    }

                    Task HandleErrorAsync(ITelegramBotClient client, Exception exception, CancellationToken cancellationToken)
                    {
                        var ErrorMessage = exception switch
                        {
                            ApiRequestException apiRequestException
                                => $"Ошибка телеграм АПИ:\n{apiRequestException.ErrorCode}\n{apiRequestException.Message}",
                            _ => exception.ToString()
                        };
                        Console.WriteLine(ErrorMessage);
                        return Task.CompletedTask;
                    }


                }
            }
        }

        private static Task HandleErrorAsync(ITelegramBotClient arg1, Exception arg2, CancellationToken arg3) => throw new NotImplementedException();
//BD
        public static SQLiteConnection DB;
        public static object MessageBox { get; private set; }

        public static void Register(string user_id, string username, string date)
        {
            try
            {
                DB = new SQLiteConnection("Data Source =userbdd.db;");
                DB.Open();

                SQLiteCommand regcmd = DB.CreateCommand();
                regcmd.CommandText = "INSERT INTO user (user_id, username , date) VALUES(@user_id, @username ,@date ) ";
                regcmd.Parameters.AddWithValue("@user_id", user_id);
                regcmd.Parameters.AddWithValue("@username", username);
                regcmd.Parameters.AddWithValue("@date", date);
                regcmd.ExecuteNonQuery();

                DB.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex);
            }

        }

 //ексель то є не фмст-_-
        public static async void ReadExcel(ITelegramBotClient bot, Update update, string late, string number)
        {
            string filePath = @"C:\Users\Віктор\source\repos\tg_prac\tg_prac\tg_prac\rozklad.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook WB = excel.Workbooks.Open(filePath);
            Worksheet wKs = (Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range cell = wKs.Cells[3, 2];
            string CellValue = cell.Value.ToString();
            int colonka;
            string rozklad = "";
            string day = "";
            int ci = 2;//для запам'ятовування стовпця для номера пари
            Console.OutputEncoding = UTF8Encoding.UTF8;
            for (colonka = 1; colonka <= 96; colonka++)
            {
                if (((wKs.Cells[2, colonka]).Value == late + " - " + number) || ((wKs.Cells[2, colonka]).Value == late + "- " + number) || ((wKs.Cells[2, colonka]).Value == late + "-" + number))
                {
                    Console.WriteLine(late + " - " + number);
                    Console.WriteLine(colonka);
                    break;
                }
            }
          
                
                rozklad = "";
                if (String.IsNullOrEmpty((wKs.Cells[2, colonka]).Value))
                {
                    if (String.IsNullOrEmpty((wKs.Cells[2, colonka + 1]).Value))
                        ci = colonka + 1;
                   
                }
                else
                {
                    rozklad += "\n" + (wKs.Cells[2, colonka]).Value;
                    Console.WriteLine(rozklad);
                }
                for (int j = 2; j <= 82; j++)
                {
                    day = "";
                    if (String.IsNullOrEmpty((wKs.Cells[j, 1]).Value))
                    {
                    }
                    else
                    {
                        day += "\n" + (wKs.Cells[j, 1]).Value;
                        Console.WriteLine(day);
                        rozklad += day + "\n";
                    }
                    if (String.IsNullOrEmpty((wKs.Cells[j, colonka]).Value))
                    {
                    }
                    else
                    {
                        rozklad += (wKs.Cells[j, ci]).Value + ". " + (wKs.Cells[j, colonka]).Value + "\n";
                    }
                    if (String.IsNullOrEmpty((wKs.Cells[j + 1, 1]).Value))
                    {
                    }
                    else
                        Console.WriteLine(rozklad);
                    if (j == 82)
                        Console.WriteLine(rozklad);
                }
                await bot.SendTextMessageAsync(update.Message.Chat.Id, rozklad);


            
        
        }

    }
}
