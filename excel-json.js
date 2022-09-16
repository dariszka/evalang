const xlsx = require("xlsx")
const fs = require("fs")

const wb = xlsx.readFile("./Data.xlsx")

const ws = wb.Sheets["Sheet1"]
const data = xlsx.utils.sheet_to_json(ws);


match()

function match() {
    let names = data.map(x => x.name);
    let spokenLs = data.map(x => x.spokenL);
    let desiredLs = data.map(x => x.desiredL);
    
    let noMatchesArray = names
    let partialMatchesArray
    let fullMatchesArray = []

    for (let i = 0; i < data.length; i++) {

        let firstLoopPerson = names[i]
        let firstLoopSpokenL = spokenLs[i]
        let firstLoopDesiredL = desiredLs[i]

            for (let i = 0; i < spokenLs.length; i++) {
                let secondLoopPerson = names[i]
                let secondLoopSpokenL = spokenLs[i]
                let secondLoopDesiredL = desiredLs[i]

                if ((firstLoopSpokenL == secondLoopDesiredL) && (firstLoopDesiredL == secondLoopSpokenL)){
                    fullMatchesArray[i] = [firstLoopPerson, secondLoopPerson]

                    noMatchesArray = noMatchesArray.filter(element => 
                        (element != firstLoopPerson)&&(element!=secondLoopPerson))
                } 
            }
    } readAllResults(fullMatchesArray, noMatchesArray)
}
    

function readAllResults(fullMatches, nonMatches) {
    fullMatches.forEach((pair) => { 	
        pair.sort()
    });
        fullMatches = fullMatches.filter(function(element) {
            return element !== undefined;
        });

    for (let i = 0; i < fullMatches.length; i++) {
        let stringPair = fullMatches[i].toString();
        let formattedStringPair = stringPair.replace(","," and ");
        fullMatches[i] = formattedStringPair
    }
    let finalFullMatches =  [...new Set(fullMatches)];

    let finalNonMatches = nonMatches.toString().replace(/,/g, '\n')
    


    displayResults(finalFullMatches, finalNonMatches)
}

function displayResults(fullMatches, nonMatches) {
    let formattedFullMatches = 'Full matches: \n' + fullMatches.toString().replace(/,/g,'\n') 
    //let formattedPartialMatches = '\n' + 'No full matches were found for: \n' + partialMatches + ' is a partial match'
    let formattedNonMatches = '\n' + 'No matches were found for: \n' + nonMatches
    
    let textResult =  formattedFullMatches + formattedNonMatches
    console.log(textResult)
}
