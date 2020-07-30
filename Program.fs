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
    @"C:\Users\recs\Desktop\COVID-19_Questionnaire (?-?-?).docx"




let sign num =
    if num > 0 then "positive"
    elif num < 0 then "negative"
    else "zero"

let main() =
    Console.WriteLine("sign 5: {0}", (sign 5))

main()