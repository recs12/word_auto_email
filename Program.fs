(*
    Send automatically an email from outlook with with a word attached
*)

open System
open System.IO
open OutlookApi

// Project steps:
//===============
//
// Rename a file
// Rename change the content of the document
// Open an email and attach the document
// Write a subject
// Write a message

// Source
let docQuestionnaireSource =
    @"C:\Users\recs\macros\COVID-19_Questionnaire (xxxx-xx-xx).docx"

// Write the date of today

// Target
let docQuestionnaireTarget (date:string) : string =
    @"C:\Users\recs\Desktop\COVID-19_Questionnaire("+date+").docx"

let getDate:string =
    System.DateTime.Today.ToString("yyyy-MM-dd")

let copyDocDesktop source target =
    File.Copy(string source, string target)

let createEmail =
    // Outlook.MailItem mailItem = (Outlook.MailItem)this.Application.CreateItem(Outlook.OlItemType.olMailItem)
    // mailItem.Subject = "Questionnaire - Covid 19"
    // mailItem.To = "someone@example.com"
    // mailItem.Body = """Bonjour
    // Merci de trouver le questionnaire en pièce jointe.
    // Cordialement
    // """
    // mailItem.Importance = Outlook.OlImportance.olImportanceLow
    // mailItem.Display(true)
    0

let main() =
    let  target = docQuestionnaireTarget getDate
    printfn "- source: %s" docQuestionnaireSource
    printfn "- target: %s" target
    copyDocDesktop docQuestionnaireSource target
    Console.WriteLine("covid questionary copied to Desktop.")

main()