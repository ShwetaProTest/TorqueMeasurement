let currentData = [];
let sheet2DataList = [];
const fileColors = {};
let currentZoom = {};

// A function to get today's date in the format "06-10-2023"
function getFormattedDate() {
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    return `${day}-${month}-${year}`;
}

function showLoading() {
    document.getElementById("loading").style.display = "block";
}

function hideLoading() {
    document.getElementById("loading").style.display = "none";
}

function reverseDataIfNeeded(data, reverse) {
    if (reverse) {
        return data.map(value => -value);
    }
    return data;
}

function loadFiles() {
    const folderElement = document.getElementById("folder");
    const directionElement = document.getElementById("direction");
    const fileSelect = document.getElementById("fileSelect");

    const folder = folderElement.value;
    const direction = directionElement.value;
    if (!folder || !direction) {
        fileSelect.innerHTML = '<option value="">-- Select File --</option>';
        return;
    }

    fetch(`/get_files_ajax?folder_name=${folder}`)
        .then(response => {
            if (!response.ok) {
                // Improved error handling
                return response.text().then(text => {
                    throw new Error(`Network response was not ok. Status: ${response.status}. Message: ${text}`);
                });
            }
            return response.json();
        })
        .then(data => {
            console.log("Files received from server:", data);

            const filteredData = data.filter(file => {
                if (file.startsWith("Ready_")) {
                    switch(direction) {
                        case "CW":
                            return file.includes("CW") && !file.includes("ACW") && !file.includes("CCW");
                        case "ACW":
                            return file.includes("ACW") || file.includes("CCW");
                        case "Sweep":
                            return file.includes("Sweep");
                        case "Behavior":
                            return file.includes("Behavior");
                        default:
                            return false;
                    }
                }
                return false;
            });

            console.log("Filtered files:", filteredData);
            fileSelect.innerHTML = ''; // Clearing previous options

            filteredData.forEach(file => {
                const option = document.createElement("option");
                option.value = file;
                option.text = getAdjustedFileName(file);
                fileSelect.appendChild(option);
            });
        })
        .catch(error => {
            console.error("Error:", error);
        });
}

function getAdjustedFileName(fullName) {
    if (fullName.indexOf("Nm") === -1) {
        fullName = fullName.replace(/\.xlsx$/i, "");
    }

    let simpleName = fullName.replace("Ready_", "");
    const nmIndex = simpleName.indexOf("Nm");

    if (nmIndex > 0 && simpleName[nmIndex - 1] === '_') {
        simpleName = simpleName.substring(0, nmIndex - 1);
    } else if (nmIndex !== -1) {
        simpleName = simpleName.substring(0, nmIndex);
    }        
    return simpleName;
}

function displayPlot() {
    return new Promise((resolve, reject) => {
        // Use Plotly.js to display the plot
        Plotly.newPlot('plotDiv', currentData, {}).then(() => {
            console.log('Plot displayed.');
            resolve();
        }).catch(error => {
            console.error('Error in plotting:', error);
            reject(error);
        });
    });
}

function displayPlotWithErrorHandling() {
    try {
        return displayPlot();
    } catch (error) {
        console.error("Error while displaying plot:", error);
        throw error; // Rethrow to be caught higher up
    }
}

function addFileToComparison() {
    showLoading();
    const btn = document.querySelector("button[type='button']");
    btn.disabled = true;

    let folder = document.getElementById("folder").value;
    let direction = document.getElementById("direction").value;
    let fileSelect = document.getElementById("fileSelect");
    let selectedFile = fileSelect.value;

    let displayName = getAdjustedFileName(selectedFile);

    if (!fileColors[displayName]) {
        fileColors[displayName] = getRandomColor();
    }

    fetch(`/get_file_data?folder=${folder}&file=${selectedFile}`)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            const li = document.createElement("li");

            const removeButton = document.createElement("button");
            removeButton.innerText = "-";
            removeButton.className = "remove-button";
            removeButton.addEventListener('click', function() {
                removeFileFromComparison(displayName);
            });

            const reverseSelect = document.createElement('select');
            reverseSelect.className = "individual-reverse";
            ["Normal", "Invert"].forEach(text => {
                const option = document.createElement("option");
                option.value = text.toLowerCase();
                option.text = text;
                reverseSelect.appendChild(option);
            });
            reverseSelect.addEventListener('change', updatePlot);
            
            // Create the subtraction input
            const subtractInput = document.createElement("input");
            subtractInput.type = "number";
            subtractInput.className = "adjustment-input";
            subtractInput.placeholder = "Correction Time";
            subtractInput.addEventListener('input', handleSubtractionChange.bind(null, displayName));
            
            li.appendChild(removeButton);
            li.appendChild(reverseSelect);
            li.appendChild(subtractInput);

            const fileNameSpan = document.createElement("span");
            fileNameSpan.innerText = " " + displayName;
            li.appendChild(fileNameSpan);

            document.getElementById("displaySelectedFilesList").appendChild(li);

            // Extract the subtraction value
            const subtractValue = parseFloat(subtractInput.value) || 0;
            data.x = data.x.map(time => time - subtractValue);

            const commonX = data.x;
            currentData.push({
                x: commonX,
                y: data.y,
                originalY: [...data.y],
                type: 'scatter',
                name: `${displayName}`,
                line: {
                    color: fileColors[displayName]
                },
                originalX: [...commonX]
            });

            if (direction === "Behavior" && data.temperature) {
                currentData.push({
                    x: commonX,
                    y: data.temperature,
                    type: 'scatter',
                    name: `${displayName}`,
                    line: {
                        color: fileColors[displayName],
                        dash: 'dot'
                    }
                });
            }

            if (direction !== "Sweep" && direction !== "Behavior") {
                return fetch(`/get_sheet2_data?folder=${folder}&file=${selectedFile}`);
            } else {
                return null;
            }
        })
        .then(response => {
            if (response && !response.ok) {
                throw new Error('Failed to fetch sheet2 data');
            }
            return response ? response.json() : null;
        })
        .then(data => {
            if (data) {
                sheet2DataList.push({ data: data, displayName: displayName });
            }
        })
        .catch(error => {
            console.error("Error:", error);
            alert("There was an error adding the file to the comparison. Please try again.");
        })
        .finally(() => {
            hideLoading();
            btn.disabled = false;
        });
}

function removeFileFromComparison(fileName) {
    // Remove the data from the currentData array
    const dataIndex = currentData.findIndex(d => d.name === fileName);
    if (dataIndex !== -1) {
        currentData.splice(dataIndex, 1);
    }

    // Remove the file's color mapping
    delete fileColors[fileName];

    // Remove the data from the sheet2DataList array
    const sheet2DataIndex = sheet2DataList.findIndex(d => d.displayName === fileName);
    if (sheet2DataIndex !== -1) {
        sheet2DataList.splice(sheet2DataIndex, 1);
    }

    // Remove the list item from the display
    const filesList = document.getElementById("displaySelectedFilesList");
    const listItem = Array.from(filesList.children).find(li => li.querySelector('span').innerText.trim() === fileName);
    if (listItem) {
        filesList.removeChild(listItem);
    }

    // Remove the column from sheet2Data table if it exists
    const sheet2Div = document.getElementById('sheet2Data');
    if (sheet2Div.querySelector("table")) {
        let table = sheet2Div.querySelector("table");
        
        // Find the header with the filename
        let headers = Array.from(table.querySelectorAll("th"));
        let headerToRemove = headers.find(th => th.innerText === fileName);
        if (headerToRemove) {
            let columnIndex = Array.from(headerToRemove.parentNode.children).indexOf(headerToRemove);
            Array.from(table.querySelectorAll("tr")).forEach(row => {
                row.removeChild(row.children[columnIndex]);
            });
        }
    }
    updatePlot();
}

function handleSubtractionChange(displayName, event) {
    console.log('Subtraction input changed');
    const subtractValue = parseFloat(event.target.value) || 0;
    const dataset = currentData.find(d => d.name === displayName);
    if (dataset) {
        dataset.x = dataset.originalX.map(time => time - subtractValue);
        Plotly.redraw('plotDiv'); // Assuming you have an updatePlot function to redraw the graph
    }
}

function getRandomColor() {
    const randomInt = (min, max) => Math.floor(Math.random() * (max - min + 1) + min);
    return `rgb(${randomInt(0, 255)}, ${randomInt(0, 255)}, ${randomInt(0, 255)})`;
}

function formatValue(value) {
    const number = parseFloat(value);
    if (!isNaN(number)) {
        return number.toFixed(2); // Round to two decimal places
    }
    return value; // If it's not a number, return the original value
}

function formatDate(inputDate) {
    // If inputDate is in full date string format, convert to Date object
    if (inputDate.includes('GMT')) {
        inputDate = new Date(inputDate);
    }
    
    let day, month, year;

    // If inputDate is a Date object, extract the components directly
    if (inputDate instanceof Date) {
        day = String(inputDate.getDate()).padStart(2, '0');
        month = String(inputDate.getMonth() + 1).padStart(2, '0');
        year = inputDate.getFullYear();
    } else {
        let parts;

        // Check for '.' separator
        if (inputDate.includes('.')) {
            parts = inputDate.split('.');
        }
        // Check for '/' separator
        else if (inputDate.includes('/')) {
            parts = inputDate.split('/');
        }
        // Check for '-' separator
        else if (inputDate.includes('-')) {
            parts = inputDate.split('-');
        }

        if (parts && parts.length === 3) {
            day = parts[0].padStart(2, '0');
            month = parts[1].padStart(2, '0');
            year = parts[2];
        }
    }

    if (day && month && year) {
        return `${day}-${month}-${year}`;
    }

    // Return inputDate if not in expected format
    return inputDate;
}

function displaySheet2Data(data, displayName) {
    const sheet2Div = document.getElementById('sheet2Data');
    const colorForThisFile = fileColors[displayName];

    const structuredData = {
        "max overshoot torque [Nm]": data.values[0],
        "max peak torque [Nm]": data.values[1],
        "slewrate [Nm/ms]": data.values[2],
        "effective slewrate [Nm/ms]": data.values[3],
        "rise time [ms]": data.values[4],
        "settling time [s]": data.values[5],
        "holding torque [Nm]": data.values[6],
        "measurement time [s]": data.values[7],
        "MAX zero error [Nm]": data.values[8],
        "MIN zero error [Nm]": data.values[9],
        "Average zero error (+-) [Nm]": data.values[10],
        "Note (TRQ reference band)": data.values[11],
        "Apply input [Index]": data.values[12],
        "Apply input [s]": data.values[13],
        "Tested by": data.values[14],
        "Tested on": data.values[15]
    };

    // Create a document fragment for efficient appending
    //const fragment = document.createDocumentFragment();
    // Helper function for formatted value assignment
    function getFormattedValue(parameter, value) {
        if (value === undefined || value === "undefined" || value === "") {
            return "Not Specified";
        } else if (parameter === "Tested on") {
            return formatDate(value);
        } else {
            return formatValue(value);
        }
    }
    
    if (!sheet2Div.querySelector("table")) {
        // If no table, create one with the basic structure
        let table = document.createElement("table");

        // Create header row
        let headerRow = document.createElement("tr");

        // Parameter Name Header
        let th1 = document.createElement("th");
        th1.innerText = "Parameters [unit]";
        headerRow.appendChild(th1);

        // File Name Header
        let th2 = document.createElement("th");
        th2.innerText = displayName;
        th2.style.color = colorForThisFile;
        headerRow.appendChild(th2);

        table.appendChild(headerRow);

        // Populate rows with parameter names and values
        for (let parameter in structuredData) {
            let tr = document.createElement("tr");

            // Parameter Name Cell
            let td1 = document.createElement("td");
            td1.innerText = parameter;
            tr.appendChild(td1);

            let td2 = document.createElement("td");

            td2.innerText = getFormattedValue(parameter, structuredData[parameter]);


            // Add a click listener to each table data cell
            td2.addEventListener("click", function() {
                console.log("Clicked:", displayName, parameter, structuredData[parameter]); // Debug log
                highlightPointOnPlot(displayName, parameter, structuredData[parameter]);
            });
            
            tr.appendChild(td2);
            table.appendChild(tr);
        }
        sheet2Div.appendChild(table);

    } else {
        // If the table already exists, add a new column for the new file
        let table = sheet2Div.querySelector("table");

        // Insert a new header cell in the first row for the new file's display name
        let newHeader = document.createElement("th");
        newHeader.innerText = displayName;
        newHeader.style.color = colorForThisFile;
        table.querySelector("tr").appendChild(newHeader);

        // Now, insert data cells for the new file in the subsequent rows
        let rows = table.querySelectorAll("tr:not(:first-child)");  // Exclude the header row
        let rowIndex = 0;

        for (let parameter in structuredData) {
            let newCell = document.createElement("td");
            newCell.innerText = getFormattedValue(parameter, structuredData[parameter]);
            //newCell.style.color = colorForThisFile;
            if (rows[rowIndex]) {
                rows[rowIndex].appendChild(newCell);
                rowIndex++;
            }
            
        }
    }
}

function highlightPointOnPlot(displayName,parameter, value) {
    // Debug logs
    console.log("Highlight function called with:", displayName, parameter, value);

    // Find the relevant data point
    let dataPoint = null;
    for (let dataset of currentData) {
        if (dataset.name === displayName) { 
            for (let i = 0; i < dataset.x.length; i++) {
                if (dataset.x[i] === parameter && dataset.y[i] === value) {
                    dataPoint = { x: dataset.x[i], y: dataset.y[i] };
                    break;
                }
            }
        }
    }

    if (dataPoint) {
        // Debug log for found point
        console.log("Data point found:", dataPoint);

        // Create a new annotation for the point
        const annotation = {
            x: dataPoint.x,
            y: dataPoint.y,
            xref: 'x',
            yref: 'y',
            text: `(${dataPoint.x}, ${dataPoint.y})`,
            showarrow: true,
            arrowhead: 4,
            ax: 0,
            ay: -40
        };

        const layout = {
            annotations: [annotation]
        };

        // Update the plot with the new annotation
        Plotly.relayout('plotDiv', layout);
    } else {
        console.error('Data point not found for the given parameter and value.');
    }
}

function getPlotName(fullName) {
    return fullName.split('/').pop().replace("Ready_", "").replace(".xlsx", "");
}

function updatePlot() {
    const reverseOptions = document.querySelectorAll('.individual-reverse');

    currentData.forEach((data, index) => {
        const invert = reverseOptions[index].value === 'invert';
        data.y = reverseDataIfNeeded([...data.originalY], invert); // Always reverse based on the original data
    });

    const layout = {
        xaxis: {
            title: 'Time [s]',
            gridcolor: '#e1e5ed',
            zerolinecolor: '#e1e5ed',
            tickfont: {
                family: 'Arial, sans-serif',
                size: 12,
                color: '#333'
            },
            titlefont: {
                family: 'Arial, sans-serif',
                size: 14,
                color: '#333'
            },
            showline: true,
            linewidth: 2,
            linecolor: '#333',
            fixedrange: false
        },
        yaxis: {
            title: 'Torque [Nm]',
            gridcolor: '#e1e5ed',
            zerolinecolor: '#e1e5ed',
            tickfont: {
                family: 'Arial, sans-serif',
                size: 12,
                color: '#333'
            },
            titlefont: {
                family: 'Arial, sans-serif',
                size: 14,
                color: '#333'
            },
            showline: true,
            linewidth: 2,
            linecolor: '#333',
            fixedrange: false
        },
        autosize: true,
        height: 500,
        paper_bgcolor: 'white',
        plot_bgcolor: 'white',
        margin: {
            l: 80,
            r: 50,
            b: 120,
            t: 10,
            pad: 4
        },
        showlegend: true,
        legend: {
            x: 0.02,
            y: 1,
            yanchor: 'bottom',
            xanchor: 'left',
            orientation: 'h'
        },
        dragmode: 'pan'  
    };
    const layoutToUpdate = Object.assign({}, layout, currentZoom);
    Plotly.react('plotDiv', currentData, layoutToUpdate);
}

function viewPlot() {
    if (currentData.length) {
        showLoading();
        setTimeout(() => {
            // Display the plot
            displayPlot().then(() => {
                // Clear any existing tables
                const sheet2Div = document.getElementById('sheet2Data');
                sheet2Div.innerHTML = "";
    
                // Loop through sheet2DataList to display all tables
                for (const data of sheet2DataList) {
                    displaySheet2Data(data.data, data.displayName);
                }
            }).catch(error => {
                console.error("An error occurred while plotting:", error);
            }).finally(() => {
                hideLoading();
            });
        }, 1000);
    } else {
        alert("Please add files to the comparison before viewing the plot.");
    }
}

function submitFormForPlot() {
    // Collect all the data using FormData and make an AJAX call to update the plot
    let formData = new FormData(document.querySelector('form'));
    fetch('/', {
        method: 'POST',
        body: formData
    }).then(response => response.text())
    .then(data => {
        var parser = new DOMParser();
        var doc = parser.parseFromString(data, 'text/html');
        var newPlotDiv = doc.querySelector('#plotDiv');
        document.getElementById('plotDiv').innerHTML = newPlotDiv.innerHTML;
    });
}

document.addEventListener("DOMContentLoaded", function() {
    document.getElementById('currentDate').textContent = getFormattedDate();
    const bgColor = getComputedStyle(document.body).backgroundColor;
    document.getElementById('logo').style.backgroundColor = bgColor;
    document.getElementById('endorLogo').style.backgroundColor = bgColor;
    document.getElementById('downloadButton').addEventListener('click', downloadAsImage);

    document.getElementById("displaySelectedFilesList").addEventListener("change", function(e) {
        if (e.target && e.target.classList.contains("individual-reverse")) {
            updatePlot();
        }
    });

    document.getElementById('plotDiv').on('plotly_relayout', function(data) {
        currentZoom = {
            'xaxis.range[0]': data['xaxis.range[0]'],
            'xaxis.range[1]': data['xaxis.range[1]'],
            'yaxis.range[0]': data['yaxis.range[0]'],
            'yaxis.range[1]': data['yaxis.range[1]']
        };
    });

    // Drag and Drop for divider and rowDivider
    const divider = document.getElementById('divider');
    const selection = document.querySelector('.selection');
    const plotDiv = document.getElementById("plotDiv");
    const sheet2Data = document.getElementById("sheet2Data");
    const rowDivider = document.getElementById("rowDivider");

    let isResizing = false;
    let startX = 0;
    let startWidth = 0;

    let isResizingRow = false;
    let startY = 0;
    let startHeight = 0;

    rowDivider.addEventListener('mousedown', (e) => {
        isResizingRow = true;
        startY = e.clientY;
        startHeight = plotDiv.offsetHeight;
        document.addEventListener('mousemove', handleRowMouseMove);
        document.addEventListener('mouseup', handleRowMouseUp);
    });

    rowDivider.addEventListener('dblclick', () => {
        plotDiv.style.height = '75vh';  
        sheet2Data.style.height = '25vh'; 
    });

    divider.addEventListener('mousedown', (e) => {
        e.preventDefault();
        isResizing = true;
        startX = e.clientX;
        startWidth = parseInt(document.defaultView.getComputedStyle(selection).width, 10);
        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);
        document.addEventListener('mouseleave', handleMouseUp);
    });

    function handleMouseMove(e) {
        if (!isResizing) return;
        const dx = e.clientX - startX;
        selection.style.flexBasis = (startWidth + dx) + "px";
    }

    function handleMouseUp() {
        isResizing = false;
        document.removeEventListener('mousemove', handleMouseMove);
        document.removeEventListener('mouseup', handleMouseUp);
        document.removeEventListener('mouseleave', handleMouseUp);
    }

    function handleRowMouseMove(e) {
        if (!isResizingRow) return;
        const dy = e.clientY - startY;
        const newHeight = startHeight + dy;
        if (newHeight > 100 && newHeight < (window.innerHeight - 200)) {  
            plotDiv.style.height = newHeight + 'px';
            sheet2Data.style.height = (window.innerHeight - newHeight) + 'px';
        }
    }

    function handleRowMouseUp() {
        isResizingRow = false;
        document.removeEventListener('mousemove', handleRowMouseMove);
        document.removeEventListener('mouseup', handleRowMouseUp);
    }

    // HammerJS-based swipe setup for plotDiv
    var hammerInstance = new Hammer(plotDiv);

    hammerInstance.get('swipe').set({ direction: Hammer.DIRECTION_VERTICAL });

    hammerInstance.on('swipeup', function() {
        console.log('Swiped up on the plot!');
    });

    hammerInstance.on('swipedown', function() {
        console.log('Swiped down on the plot!');
    });
});

function simpleDownload() {
    html2canvas(document.querySelector("body")).then(canvas => {
        let link = document.createElement('a');
        link.download = 'Torque_measurement.jpg'; 
        link.href = canvas.toDataURL('image/jpeg', 1.0); 
        link.click();
    });
}

function downloadAsImage() {
    const container = document.createElement('div');
    const plotClone = document.querySelector('.plot').cloneNode(true);
    container.appendChild(plotClone);

    const table = document.querySelector('#sheet2Data').querySelector('table');
    if (table) {
        const tableClone = table.cloneNode(true);
        container.appendChild(tableClone);
    }

    document.body.appendChild(container);
    container.style.display = 'none';

    html2canvas(container).then(canvas => {
        canvas.toBlob(function(blob) {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'screenshot.jpg';
            link.click();

            setTimeout(() => {
                URL.revokeObjectURL(link.href);
            }, 100);

            document.body.removeChild(container);
        }, 'image/jpeg', 1.0);
    });
}
