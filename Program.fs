(*
    Send automatically an email from outlook with with a word attached
*)

open System
open System.IO // File.Move(source, destination)


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
let docQuestionnaireTarget =
    @"C:\Users\recs\Desktop\COVID-19_Questionnaire (xxxx-xx-xx).docx"


let moveDocDesktop source target =
    File.Move(string source, string target)



let main() =
    moveDocDesktop docQuestionnaireSource docQuestionnaireTarget
    Console.WriteLine("Done")

main()