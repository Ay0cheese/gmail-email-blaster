open System
open System.Net.Mail
// Learn more about F# at http://docs.microsoft.com/dotnet/fsharp

open System
open FSharp.Data


type EmailData = CsvProvider<"monsterText.csv">
let sendMailMessage sender email subject body template =
    
    let port = 587
    let password = "ifnv sgnx qxpk qfbr"
    // ujac wbvo fhlr awdq
    // let password = "******" // ConfigurationManager.AppSettings.["mailpassword"] |> my-decrypt
    let server = "smtp.gmail.com" // ConfigurationManager.AppSettings.["mailserver"]

    let msg = new MailMessage( sender, email, subject, body )
    let format = SmtpDeliveryFormat()
    msg.IsBodyHtml <- true
    // printfn "Body is HTML: %b" (msg.IsBodyHtml) 
    let client = new SmtpClient(server, port, EnableSsl = true)
    client.EnableSsl <- true
    client.DeliveryMethod <- SmtpDeliveryMethod.Network
    // client.DeliveryFormat <- SmtpDeliveryFormat()
    client.Credentials <- System.Net.NetworkCredential(sender, password)
    client.SendCompleted |> Observable.add(fun e -> 
        let msg = e.UserState :?> MailMessage
        if e.Cancelled then
            ("Mail message cancelled:\r\n" + msg.Subject) |> Console.WriteLine
        if e.Error <> null then
            ("Sending mail failed for message:\r\n" + msg.Subject + 
                ", reason:\r\n" + e.Error.ToString()) |> Console.WriteLine
        if msg<>Unchecked.defaultof<MailMessage> then msg.Dispose()
        if client<>Unchecked.defaultof<SmtpClient> then client.Dispose()
    )
    // Maybe some System.Threading.Thread.Sleep to prevent mail-server hammering
    client.Send(msg)

[<EntryPoint>]
let main argv =
    let sender = "wade@monsterreviews.io" // ConfigurationManager.AppSettings.["mailsender"]

    match argv with
    | [|subject|] ->

        let subject = subject
        let fullPath = System.IO.Directory.GetCurrentDirectory() + "/" + "monsterText.csv"
        let templatePath = System.IO.Directory.GetCurrentDirectory() + "/" + "template.html"
        let template = System.IO.File.ReadAllText(templatePath)

        printfn "File to email: %s" (fullPath)
        printfn "With template: %s" (templatePath)
        printfn "Email subject: %s" (subject)
        
        let emailData = EmailData.Load(fullPath)
        emailData.Rows
        |> Seq.map ( fun lead -> 
            {|
                Name = lead.Name
                Email = lead.Email
            |}
        )
        |> Seq.iteri (
            fun currentIndex lead ->
                async {
                    if (currentIndex >= 60 && currentIndex % 60 = 0)
                    then
                        printfn "Sleeping for 5 minutes..."
                        let! a = Async.Sleep 300000
                        a
                    printfn "Sending email to: %s" lead.Email
                    let emailBody = template.Replace("{{name}}", lead.Name)
                    sendMailMessage sender lead.Email subject emailBody template
                } |> Async.RunSynchronously
        )

        printfn "Done sending emails!!"

    | _ -> failwith "1 argument required!" 

    0
