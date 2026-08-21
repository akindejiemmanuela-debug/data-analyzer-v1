console.log("dashboard_charts.js LOADED");
const chart = echarts.init(document.getElementById("chart"));
chart.on("click", function(params){

    window.location.href =
        "/drilldown?category=" +
        encodeURIComponent(params.name) +
        "&x_column=" +
        encodeURIComponent(xColumn) +
        "&y_column=" +
        encodeURIComponent(yColumn);

});
let option = {

    title: {

        animationDuration: 5000,
        animationEasing: "cubicOut",
        text: bar_title,
        left: "center"
    },

    tooltip: {
        trigger: "axis"
    },

    xAxis: {
        type: "category",
        data: labels,
        axisLabel: {
            rotate: 45
        }
    },

    yAxis: {
        type: "value"
    },

    series: [{
    type: "bar",
    data: values,

    itemStyle: {
        color: function(params){

            const colors = [
                "#4e73df",
                "#1cc88a",
                "#36b9cc",
                "#f6c23e",
                "#e74a3b",
                "#6f42c1",
                "#fd7e14",
                "#20c997"
            ];

            return colors[params.dataIndex % colors.length];
        }
    }
}]
};

chart.setOption(option);
const pieChart = echarts.init(document.getElementById("pieChart"));

let pieOption = {
    animationDuration: 5000,
    animationEasing: "cubicOut",
    title: {
        text: pie_title,
        left: "center"
    },

    tooltip: {},

    series: [{
        color: [
"#4e73df",
"#1cc88a",
"#36b9cc",
"#f6c23e",
"#e74a3b",
"#6f42c1",
"#fd7e14",
"#20c997"
],
        type: "pie",

        radius: "70%",

        data: labels.map((label, i) => ({
            name: label,
            value: values[i]
        }))

    }]

};

pieChart.setOption(pieOption);
const lineChart = echarts.init(document.getElementById("lineChart"));

const histChart = echarts.init(document.getElementById("histChart"));

const scatterChart = echarts.init(document.getElementById("scatterChart"));

let lineOption = {
    animationDuration: 5000,
    animationEasing: "cubicOut",

    title: {
        text: line_title,
        left: "center"
    },

    xAxis: {
        type: "category",
        data: labels
    },

    yAxis: {
        type: "value"
    },

    series: [{
        type: "line",
        data: values,
        smooth: true,

        itemStyle: {
            color: "#4e73df"
        }
    }]
};
lineChart.setOption(lineOption);


// FIX: Number("") is 0, not NaN — that was silently turning blank cells
// into real zero values in the histogram/scatter data. Strip blanks first.
let numericValues = values
    .filter(v => v !== "" && v !== null && v !== undefined)
    .map(Number)
    .filter(v => !isNaN(v));

let histOption = {
    animationDuration: 5000,
    animationEasing: "cubicOut",

    title: {
        text: hist_title,
        left: "center"
    },

    tooltip: {
    trigger: "axis",

    formatter: function(params) {

        const item = params[0];

        return `
            <b>Value Range:</b> ${item.name}<br>
            <b>Frequency:</b> ${item.value}
        `;
    }
},

    xAxis: {
        type: "category",
        // FIX: this was "y_column" (snake_case), an identifier that is
        // never declared anywhere — the inline <script> block in your
        // template only defines "yColumn" (camelCase). Referencing an
        // undefined global here threw a ReferenceError while building
        // histOption, which stopped the script before histChart.setOption()
        // ran, before scatterOption was ever built, and before the resize
        // listener below got registered. That's why both the histogram and
        // the scatter chart were blank.
        name: yColumn,
        data: []
    },

    yAxis: {
        type: "value",
        name: "Frequency"
    },

    series: [{
        type: "bar",
        data: [],
        barWidth: "90%",

        itemStyle: {
            color: "#1cc88a"
        }
    }]
};
let bins = [];
let frequencies = [];

if (numericValues.length > 0) {

    let minValue = Math.min(...numericValues);
    let maxValue = Math.max(...numericValues);

    // If all values are the same
    if (minValue === maxValue) {

        bins.push(minValue.toString());
        frequencies.push(numericValues.length);

    } else {

        let binCount = 10;
        let binWidth = (maxValue - minValue) / binCount;

        for (let i = 0; i < binCount; i++) {

            let start = minValue + (i * binWidth);
            let end = start + binWidth;

            let count = numericValues.filter(v => {

                if (i === binCount - 1) {
                    return v >= start && v <= end;
                }

                return v >= start && v < end;

            }).length;

            bins.push(
                `${start.toFixed(1)} - ${end.toFixed(1)}`
            );

            frequencies.push(count);
        }
    }
}

histOption.xAxis.data = bins;
histOption.series[0].data = frequencies;


histChart.setOption(histOption);
console.log("Histogram bins:", bins);
console.log("Histogram frequencies:", frequencies);
console.log("Numeric values:", numericValues);

let scatterData = numericValues.map((v, i) => [i + 1, v]);

let scatterOption = {
    animationDuration: 5000,
    animationEasing: "cubicOut",

    title: {
        text: scatter_title,
        left: "center"
    },

    tooltip: {
        trigger: "item"
    },

    xAxis: {
        type: "value",
        name: "Record"
    },

    yAxis: {
        type: "value",
        name: yColumn
    },

    series: [{
        type: "scatter",
        symbolSize: 10,
        data: scatterData,

        itemStyle: {
            color: "#f6c23e"
        }
    }]
};



scatterChart.setOption(scatterOption);
console.log("Scatter values:", values);
function filterChart(){

    let value = document.getElementById("chartFilter").value;

    if(value==""){

        chart.setOption({
            xAxis:{data:labels},
            series:[{data:values}]
        });

        pieChart.setOption({
            series:[{
                data:labels.map((l,i)=>({
                    name:l,
                    value:values[i]
                }))
            }]
        });

        lineChart.setOption({
            xAxis:{data:labels},
            series:[{data:values}]
        });

        histChart.setOption({
    xAxis: {
        data: bins
    },
    series: [{
        data: frequencies
    }]
});

        scatterChart.setOption({
            series:[{
                data:values.map((v,i)=>[i+1,v])
            }]
        });

        return;
    }

    let index = labels.indexOf(value);

    let label = labels[index];

    let valueData = values[index];

    chart.setOption({
        xAxis:{data:[label]},
        series:[{data:[valueData]}]
    });

    pieChart.setOption({
        series:[{
            data:[
                {
                    name:label,
                    value:valueData
                }
            ]
        }]
    });

    lineChart.setOption({
        xAxis:{data:[label]},
        series:[{data:[valueData]}]
    });

    // Histogram
histChart.setOption({
    xAxis:{
        data:[valueData.toString()]
    },
    series:[{
        data:[1]
    }]
});

    // FIX: coerce to Number so a non-numeric valueData doesn't get plotted
    // as a string y-value on a numeric axis.
    scatterChart.setOption({
        series:[{
            data:[[1, Number(valueData)]]
        }]
    });

}


window.addEventListener("resize", function(){

    chart.resize();
    pieChart.resize();
    lineChart.resize();
    histChart.resize();
    scatterChart.resize();

});


function exportChartsToPDF() {

    console.log("===== PDF EXPORT STARTED =====");

    // Make sure charts are properly rendered
    chart.resize();
    pieChart.resize();
    lineChart.resize();
    histChart.resize();
    scatterChart.resize();

    // Get chart images
    const barImage = chart.getDataURL({
        type: "png",
        pixelRatio: 2,
        backgroundColor: "#ffffff"
    });

    const pieImage = pieChart.getDataURL({
        type: "png",
        pixelRatio: 2,
        backgroundColor: "#ffffff"
    });

    const lineImage = lineChart.getDataURL({
        type: "png",
        pixelRatio: 2,
        backgroundColor: "#ffffff"
    });

    const histogramImage = histChart.getDataURL({
        type: "png",
        pixelRatio: 2,
        backgroundColor: "#ffffff"
    });

    const scatterImage = scatterChart.getDataURL({
        type: "png",
        pixelRatio: 2,
        backgroundColor: "#ffffff"
    });


    console.log("BAR IMAGE:", barImage.length);
    console.log("PIE IMAGE:", pieImage.length);
    console.log("LINE IMAGE:", lineImage.length);
    console.log("HISTOGRAM IMAGE:", histogramImage.length);
    console.log("SCATTER IMAGE:", scatterImage.length);


    // Create POST form
    const form = document.createElement("form");

    form.method = "POST";
    form.action = "/export_pdf";


    function addField(name, value) {

        const input = document.createElement("input");

        input.type = "hidden";
        input.name = name;
        input.value = value;

        form.appendChild(input);

    }


    addField("bar", barImage);

    addField("pie", pieImage);

    addField("line", lineImage);

    addField("histogram", histogramImage);

    addField("scatter", scatterImage);


    // Send workspace ID
    addField(
        "workspace_id",
        "{{ session.get('workspace_id', '') }}"
    );


    document.body.appendChild(form);


    console.log("Submitting /export_pdf...");

    form.submit();

}