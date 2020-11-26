import os
import random

from pandas import DataFrame
from openpyxl import load_workbook
from pyproj import Transformer

import folium
from folium import plugins
from branca.element import Template, MacroElement


def draw_map(road_name, degree, scale):
    load_wb = load_workbook(os.path.abspath("가로수좌표/" + road_name + ".xlsx"), data_only=True)
    load_ws = load_wb['Sheet1']
    color_list = ['rgba(255,0,0,0.5)', 'rgba(255,94,0,0.5)', 'rgba(255,228,0,0.5)', 'rgba(0,255,0,0.5)',
                  'rgba(0,0,255,0.5)', 'rgba(50,0,155,0.5)', 'rgba(255,0,204,0.5)', 'rgba(153,51,102, 0.5)', 
                  'rgba(51,255,255,0.5)', 'rgba(0,0,0,0.5)']

    color_dic = {}

    all_values = []
    for row in load_ws.rows:
        row_value = []
        for cell in row:
            row_value.append(cell.value)
        all_values.append(row_value)

    ex = {'경도': [],
          '위도': [],
          '구분': [],
          '도로': []}

    for row in all_values:
        transformer = Transformer.from_crs('epsg:5186', 'epsg:4326')
        x, y = transformer.transform(row[2], row[1])
        ex['경도'].append(y)
        ex['위도'].append(x)
        ex['구분'].append(str(row[0]))
        ex['도로'].append(row[3])

    road_list = set(ex['도로'])
    print(sorted(road_list))

    
    for idx, val in enumerate(sorted(road_list)):
        color_dic[val] = color_list[idx]

    ex = DataFrame(ex)
    print(ex)

    # 지도의 중심을 지정하기 위해 위도와 경도의 평균 구하기
    lat = ex['위도'].mean()
    long = ex['경도'].mean()

    print(lat)
    print(long)

    # 지도 띄우기
    m = folium.Map([lat, long], zoom_start=18)

    for i in ex.index:
        sub_lat = ex.loc[i, '위도']
        sub_long = ex.loc[i, '경도']

        # stranger = False
        title = ''
        popup = ''
        if road_name != ex.loc[i, '도로']:
            popup += '<div style="transform: scale(0.55);">'
            popup += ex.loc[i, '도로'] + '</div>'
            # stranger = True
        title += '<div style="transform: scale(' + str(scale) + ');">' + ex.loc[i, '구분'] + '</div>'
        popup += title

        # color = 'red'
        # if stranger:
        #     color = 'green'
        # bg_color = ''
        # if color == 'red':
        #     bg_color = 'rgba(255,0,0,0.3)'
        # elif color == 'green':
        #     bg_color = 'rgba(0,255,0,0.3)'

        icon_number = plugins.BeautifyIcon(
            border_color=color_dic[ex.loc[i, '도로']],
            iconSize=[5, 5],
            iconAnchor=[5, 5],
            backgroundColor=color_dic[ex.loc[i, '도로']],
            text_color='rgba(0,0,0,1)',
            customClasses='myAnchor',
            border_width=0.2,
            number=title,
            inner_icon_style='margin-top:0px;'
                             'margin-left:0px;'
                             'transform: rotate(' + str(degree) + 'deg);'
                                                                  'text-shadow: -0.1px 0 white, 0 0.1px white, 0.1px 0 white, 0 -0.1px white;'
                                                                  'font-size: 10px;'
                                                                  'font-weight: bold;'
                                                                  'white-space: nowrap;'
                                                                  'line-height: 0.7;'
        )
        folium.Marker(location=[sub_lat, sub_long], popup=popup, icon=icon_number).add_to(m)
    ###################범례 추가######################
    template = """
    {% macro html(this, kwargs) %}

    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>jQuery UI Draggable - Default functionality</title>
      <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">

      <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
      <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

      <script>
      $( function() {
        $( "#maplegend" ).draggable({
                        start: function (event, ui) {
                            $(this).css({
                                right: "auto",
                                top: "auto",
                                bottom: "auto"
                            });
                        }
                    });
    });

      </script>
    </head>
    <body>


    <div id='maplegend' class='maplegend' 
        style='position: absolute; z-index:9999; border:2px solid grey; background-color:rgba(255, 255, 255, 0.8);
         border-radius:6px; padding: 10px; font-size:14px; right: 20px; bottom: 20px;'>

    <div class='legend-title'>범례</div>
    <div class='legend-scale'>
      <ul class='legend-labels'>
    """

    for val in road_list:
        template += "<li><span style='background:"+color_dic[val]+";'></span>"+val+"</li>"

    template += """
      </ul>
    </div>
    </div>

    </body>
    </html>

    <style type='text/css'>
      .maplegend .legend-title {
        text-align: left;
        margin-bottom: 5px;
        font-weight: bold;
        font-size: 90%;
        }
      .maplegend .legend-scale ul {
        margin: 0;
        margin-bottom: 5px;
        padding: 0;
        float: left;
        list-style: none;
        }
      .maplegend .legend-scale ul li {
        font-size: 80%;
        list-style: none;
        margin-left: 0;
        line-height: 18px;
        margin-bottom: 2px;
        }
      .maplegend ul.legend-labels li span {
        display: block;
        float: left;
        height: 16px;
        width: 30px;
        margin-right: 5px;
        margin-left: 0;
        border: 1px solid #999;
        }
      .maplegend .legend-source {
        font-size: 80%;
        color: #777;
        clear: both;
        }
      .maplegend a {
        color: #777;
        }
    </style>
    {% endmacro %}"""

    macro = MacroElement()
    macro._template = Template(template)

    m.get_root().add_child(macro)

    m.save("html/" + road_name + '.html')
