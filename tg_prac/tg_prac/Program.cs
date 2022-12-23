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




namespace tg_bot
{
    class Program
    {
        static TelegramBotClient botClient = new TelegramBotClient("5802121972:AAGc5CvkTyuDftTi172ACwurBtom9F3nS_A");
        static void Main(string[] args)
        {
            Console.WriteLine($"Бот {botClient.GetMeAsync().Result.FirstName} запустився.");

            using var cts = new CancellationTokenSource();
            var cancellationToken = cts.Token;
            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = { }
            };

            botClient.StartReceiving(
             UpdateHandler,
            HandleMessage,
             receiverOptions,
             cancellationToken);



    Console.ReadLine();

            cts.Cancel();

            async Task HandleUpdatesAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
            {
                if (update.Type == UpdateType.Message && update?.Message?.Text != null)
                {
                    await HandleMessage(botClient, update.Message);
                    return;
                }

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

                        //excel
                        if (message.Text == "/exc")
                        {
                            readExcel();
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





                        await botClient.SendTextMessageAsync(message.Chat.Id, $"You said:\n{message.Text}");
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

        private static Task HandleMessage(ITelegramBotClient botClient, Message message)
        {
            throw new NotImplementedException();
        }

        private static Task HandleMessage(ITelegramBotClient arg1, Exception arg2, CancellationToken arg3)
        {
            throw new NotImplementedException();
        }


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

        public static void readExcel()
        {
            string filePath = @"C:\Users\Віктор\source\repos\tg_prac\tg_prac\tg_prac\rozklad.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet wks;
            wb = excel.Workbooks.Open(filePath);
            wks = wb.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range cell = wks.Cells[3, 2];
            string CellValue = cell.Value.ToString();
            Console.WriteLine(CellValue);
        }


    }
}