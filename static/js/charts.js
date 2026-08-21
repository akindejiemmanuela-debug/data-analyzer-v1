
  
const myChart=echarts.init(document.getElementById("chart"));  
const myChart2 = echarts.init(document.getElementById("chart2"));  
let option={  
  
animation:true,  
  
animationDuration:3000,  
  
animationDelay:function(idx){  
    return idx*40;  
},  
  
animationDurationUpdate:2000,  
  
animationEasing:"bounceOut",  
  
animationEasingUpdate:"cubicOut",  
  
tooltip:{  
  
trigger:"axis",  
  
axisPointer:{  
type:"shadow"  
}  
  
},  
toolbox:{  
  
right:20,  
  
feature:{  
  
saveAsImage:{},  
  
restore:{},  
  
dataZoom:{},  
  
dataView:{readOnly:true}  
  
}  
  
},  
dataZoom:[  
  
{  
  
type:"inside"  
  
},  
  
{  
  
type:"slider",  
  
start:0,  
  
end:30  
  
}  
  
],  
title:{

    text: yTitle + " Analysis",

    left:"center"

},
  
xAxis:{  
  
type:"category",  
  
data:labels,  
  
axisLabel:{  
  
rotate:45,  
  
interval:0,  
  
fontSize:12  
  
},  
  
axisTick:{  
  
alignWithLabel:true  
  
}  
  
},  
yAxis:{  
  
type:"value",  
  
splitLine:{  
  
show:true  
  
}  
  
}, 
series:[]
};

let option2 = JSON.parse(JSON.stringify(option));    
if(chartType=="bar"){  
  
option.series=[{

    type:"bar",

    data:values,

    barWidth:"55%",

    universalTransition:true,

    itemStyle:{
        color:function(params){

            const colors=[
                "#3b82f6",
                "#22c55e",
                "#f59e0b",
                "#ef4444",
                "#8b5cf6",
                "#06b6d4",
                "#ec4899",
                "#84cc16",
                "#ff9800",
                "#795548"
            ];

            return colors[params.dataIndex % colors.length];
        },

        borderRadius:[8,8,0,0]
    },

    emphasis:{
        scale:true,
        focus:"series"
    },

    label:{
        show:true,
        position:"top"
    }

}];
}  
else if(chartType=="line"){  
  
option.series=[{  
  
type:"line",  
  
smooth:true,  
  
data:values,  
  
symbol:"circle",  
  
symbolSize:10,  
  
lineStyle:{  
  
width:4  
  
},  
  
areaStyle:{},  
  
itemStyle:{  
  
color:"#1565C0"  
  
}  
  
}];  
  
}  
  
else if(chartType=="scatter"){  
  
option.series=[{  
  
type:"scatter",  

data: values.map((v,i)=>[i + 1, v]), 
  
symbolSize:15,  
  
itemStyle:{  
  
color:"#FF5722"  
  
}  
  
}];  
  
}  
  
else if(chartType=="pie"){  
  
option={  
  
animation:true,  
  
tooltip:{},  
  
toolbox:{  
  
feature:{  
  
saveAsImage:{}  
  
}  
  
},  
  
title:{

    text: yTitle + " Distribution",

    left:"center"

},
  
series:[{  
  
type:"pie",  
  
radius:["45%","75%"],  
  
data:labels.map((l,i)=>({  
  
name:l,  
  
value:values[i]  
  
})),  
  
emphasis:{  
  
itemStyle:{  
  
shadowBlur:20,  
  
shadowOffsetX:0  
  
}  
  
}  
  
}]  
  
};  
  
}  

else if(chartType=="histogram"){

    option.xAxis = {
        type: "category",
        data: labels
    };

    option.yAxis = {
        type: "value"
    };

    option.series = [{
        type: "bar",
        data: values
    }];

} 
  
else if(chartType=="box"){  
  
option={  
  
title:{  
  
text:"Box Plot"  
  
},  
  
dataset:[{  
  
source:[values]  
  
}],  
  
xAxis:{type:"category"},  
  
yAxis:{type:"value"},  
  
series:[{  
  
type:"boxplot"  
  
}]  
  
};  
  
}  
  
window.onload = function(){

    loadChart();

    myChart.on("click", function(params){

        let category = params.name;

        window.location.href =
            "/drilldown?category=" + encodeURIComponent(category);

    });

};
setTimeout(function(){  
  
myChart.dispatchAction({  
  
type:"highlight",  
  
seriesIndex:0,  
  
dataIndex:0  
  
});  
  
},500);  
window.addEventListener("resize", () => {
    myChart.resize();
    myChart2.resize();
});
 document.querySelectorAll(".counter").forEach(counter=>{

    let target = Number(counter.innerText);
    let count = 0;

    let speed = target / 100;

    let update = ()=>{

        count += speed;

        if(count < target){

            counter.innerText = Math.floor(count);

            requestAnimationFrame(update);

        }else{

            counter.innerText = target;

        }

    };

    update();

}); 

function highlightMaximum(){

    let max = Math.max(...values);

    let index = values.indexOf(max);

    myChart.dispatchAction({
        type:"highlight",
        seriesIndex:0,
        dataIndex:index
    });

}

function highlightMinimum(){

    let min = Math.min(...values);

    let index = values.indexOf(min);

    myChart.dispatchAction({
        type:"highlight",
        seriesIndex:0,
        dataIndex:index
    });

}

function toggleTheme(){

    document.body.classList.toggle("dark-mode");

    let btn = document.getElementById("themeBtn");

    if(document.body.classList.contains("dark-mode")){
        btn.innerHTML="☀️ Light Mode";
    }else{
        btn.innerHTML="🌙 Dark Mode";
    }

}

function goDrilldown(){

    let category = labels[0];   // Default category

    window.location.href =
        "/drilldown?category=" + encodeURIComponent(category);

}

function filterChart(){

    let value = document.getElementById("chartFilter").value;

    if(value==""){

        myChart.setOption({
            xAxis:{data:labels},
            series:[{data:values}]
        });

        return;
    }

    let index = labels.indexOf(value);

    myChart.setOption({

        xAxis:{
            data:[labels[index]]
        },

        series:[{
            data:[values[index]]
        }]

    });

}

function loadChart(){

    myChart.clear();
    myChart.setOption(option, true);

    myChart2.clear();
    option2.title = {
    text: "Pie Chart",
    left: "center"
};

option2.series = [{
    type: "pie",
    radius: ["45%", "75%"],
    data: labels.map((label, i) => ({
        name: label,
        value: values[i]
    }))
}];

delete option2.xAxis;
delete option2.yAxis;

myChart2.setOption(option2, true);

}

async function refreshChart(){

    let response = await fetch("/chart_data");

    let data = await response.json();

    option.xAxis.data = data.labels;

    option.series[0].data = data.values;

    loadChart();

}

setInterval(refreshChart, 10000);
