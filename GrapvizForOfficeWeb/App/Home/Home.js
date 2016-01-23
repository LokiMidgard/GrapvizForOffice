/// <reference path="../App.js" />
/// <reference path="http://gabelerner.github.io/canvg/canvg.js" />

(function () {
    "use strict";

    // Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#set-data-to-selection').click(writeData);



            // Configure Fabric NavBar
            $('.ms-NavBar').NavBar();
            $('#teamBuilder').click(function showTb() {
                $('#tbPanel').show();
                $('#sowPanel').hide();
            });
            $('#docGen').click(function showDg() {
                $('#tbPanel').hide();
                $('#sowPanel').show();
            });
            $('#docGen').trigger("click");



            var editor = ace.edit("editor");
            editor.getSession().setMode("ace/mode/dot");
            editor.setOptions({
                maxLines: Infinity
            });

            var parser = new DOMParser();
            var worker;
            var result;
            var imageData;

            function updateGraph() {
                if (worker) {
                    worker.terminate();
                }

                document.querySelector("#output").classList.add("working");
                document.querySelector("#output").classList.remove("error");

                worker = new Worker("./worker.js");

                worker.onmessage = function (e) {
                    document.querySelector("#output").classList.remove("working");
                    document.querySelector("#output").classList.remove("error");

                    result = e.data;

                    updateOutput();
                }

                worker.onerror = function (e) {
                    document.querySelector("#output").classList.remove("working");
                    document.querySelector("#output").classList.add("error");

                    var message = e.message === undefined ? "An error occurred while processing the graph input." : e.message;

                    var error = document.querySelector("#error");
                    while (error.firstChild) {
                        error.removeChild(error.firstChild);
                    }

                    document.querySelector("#error").appendChild(document.createTextNode(message));

                    console.error(e);
                    e.preventDefault();
                }

                var params = {
                    src: editor.getSession().getDocument().getValue(),
                    options: {
                        engine: 'dot',
                        format: 'svg'
                    }
                };


                worker.postMessage(params);
            }

            function updateOutput() {
                var graph = document.querySelector("#output");

                var svg = graph.querySelector("svg");
                if (svg) {
                    graph.removeChild(svg);
                }

                var text = graph.querySelector("#text");
                if (text) {
                    graph.removeChild(text);
                }

                var img = graph.querySelector("img");
                if (img) {
                    graph.removeChild(img);
                }

                if (!result) {
                    return;
                }

                //if (document.querySelector("#format select").value == "svg" && !document.querySelector("#raw input").checked) {
                canvg('canvas', result);
                var img = canvas.toDataURL("image/png");
                imageData = img.substr("data:image/png;base64,".length);
                var svg = parser.parseFromString(result, "image/svg+xml");
                graph.appendChild(svg.documentElement);
                //} else if (document.querySelector("#format select").value == "png-image-element") {




                //    var imgElement = document.createElement("img");
                //    imgElement.src = img;
                //    graph.appendChild(imgElement);
                //} else {
                //    var text = document.createElement("div");
                //    text.id = "text";
                //    text.appendChild(document.createTextNode(result));
                //    graph.appendChild(text);
                //}
            }

            editor.on("change", function () {
                updateGraph();
            });

            //document.querySelector("#engine select").addEventListener("change", function () {
            //    updateGraph();
            //});

            //document.querySelector("#format select").addEventListener("change", function () {
            //    if (document.querySelector("#format select").value === "svg") {
            //        document.querySelector("#raw").classList.remove("disabled");
            //        document.querySelector("#raw input").disabled = false;
            //    } else {
            //        document.querySelector("#raw").classList.add("disabled");
            //        document.querySelector("#raw input").disabled = true;
            //    }

            //    updateGraph();
            //});

            //document.querySelector("#raw input").addEventListener("change", function () {
            //    updateOutput();
            //});

            updateGraph();


            // Liest Daten aus der aktuellen Dokumentauswahl und zeigt eine Benachrichtigung an
            function getDataFromSelection() {


                Word.run(function (context) {

                    var range = context.document.getSelection();
                    var paragraphs = range.paragraphs;
                    context.load(paragraphs);

                    return context.sync().then(function () {

                        var pictures = paragraphs.items[0].inlinePictures;
                        context.load(pictures);

                        return context.sync().then(function () {
                            var picture = pictures.items[0];
                            
                            editor.getSession().getDocument().setValue(picture.altTextDescription);
                        });
                    });

                })
            }
            function writeData() {


                Word.run(function (context) {

                    var range = context.document.getSelection();
                    var mybase64 = imageData;
                    var picture = range.insertInlinePictureFromBase64(mybase64, Word.InsertLocation.end);
                    picture.altTextDescription = editor.getSession().getDocument().getValue();

                })

            }

        });
    };




})();