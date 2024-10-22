var watermark_img;
var watermark2_img;
var canvas;
var canvasWidth;
var canvasHeight;
var img;

// Wait for the Office.js library to be loaded and ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office.js is ready, initializing the add-in");

        myPreloadSetup();
        initializeAddIn();
    } else {
        console.log("This add-in is not designed for the current host:", info.host);
    }
});
function myPreloadSetup() {
    console.log("'preLoad();' run");
    watermark_img = loadImage("https://sebastianrix.github.io/Hosting-Outlook_Addin_Prototype/assets/watermark.png");
    watermark2_img = loadImage("https://sebastianrix.github.io/Hosting-Outlook_Addin_Prototype/assets/icon-80.png");

    console.log("'setup();' run");
    canvasWidth = 526;
    canvasHeight = 785;
    canvas = createCanvas(canvasWidth, canvasHeight);
    const appContainer = document.getElementById("app-container");
    appContainer.appendChild(canvas.elt);
    canvas.elt.style.display = "block";
    canvas.elt.style.margin = "auto";
}


function WriteText() {
   

    console.log("'WriteText();' run");
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = result.value;
            Generate(emailBody);
        }
    });
}

function Generate(takenText) {
    console.log("'Generate();' run");
    const myText = takenText;
    const lineBreakRegex = /\r\n|\r|\n/g;
    const lines = myText.split(lineBreakRegex);

    const tempCanvas = document.createElement("canvas");
    const tempCtx = tempCanvas.getContext("2d");
    tempCtx.font = "16px Arial";

    let maxWidth = 0;
    let maxHeight = 0;
    for (const line of lines) {
        const lineWidth = tempCtx.measureText(line).width;
        maxWidth = Math.max(maxWidth, lineWidth);
        maxHeight += 20;
    }

    const padding = 20;
    const canvasWidthMail = maxWidth + padding * 2;
    const canvasHeightMail = maxHeight + padding;

    resizeCanvas(canvasWidthMail, canvasHeightMail);
    background(255, 255, 255);
    tint(255, 127);

    image(watermark2_img, 0, 0);
    textSize(16);
    fill(0);

    let y = padding;
    for (const line of lines) {
        text(line, padding, y);
        y += 20;
    }

    convertToImage(canvasWidthMail, canvasHeightMail);
}

function convertToImage(canvasWidthMail, canvasHeightMail) {
    console.log("'convertToImage();' run");
    const canvasElement = canvas.elt;
    const imageDataUrl = canvasElement.toDataURL();
    Office.context.mailbox.item.body.setAsync(
        `<img src="${imageDataUrl}" width="${canvasWidthMail}" height="${canvasHeightMail}">`,
        { coercionType: Office.CoercionType.Html },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Error setting email body: ' + result.error.message);
            } else {
                console.log('Email body set with image.');
            }
        }
    );
}

function initializeAddIn() {
    console.log("Initializing the add-in");
    const helloButton = document.getElementById("helloButton");
    if (helloButton) {
        helloButton.onclick = () => {
            WriteText();
        };
    } else {
        console.error("Could not find the 'helloButton' element");
    }
}
