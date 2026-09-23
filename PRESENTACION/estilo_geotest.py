"""Estilos compartidos para las pantallas NiceGUI de Geotest."""

from nicegui import ui


def aplicar_estilo_geotest():
    ui.colors(primary='#5979E6', secondary='#253B83', accent='#E23742')
    ui.add_head_html('''
    <style>
      body { background: #F8FAFF; color: #202938; font-family: Arial, sans-serif; }
      .geotest-header { background: #fff; color: #202938; border-bottom: 1px solid #DFE4EF; box-shadow: none; }
      .geotest-sidebar { background: #151E39; color: #fff; }
      .geotest-sidebar .q-item { border-radius: 10px; margin: 4px 10px; }
      .geotest-sidebar .q-item:hover { background: #303B59; }
      .geotest-card { background: #fff; border: 1px solid #DFE4EF; border-radius: 17px; box-shadow: none; }
      .geotest-banner {
            overflow: hidden;
            padding-left: 36px !important;
            background-color: #fff;
            background-image: linear-gradient(to bottom,
            #3045B4 0%, #3045B4 33.333%,
            #95A9EF 33.333%, #95A9EF 66.666%,
            #E23742 66.666%, #E23742 100%);
            background-position: left center;
            background-size: 5px 100%;
            background-repeat: no-repeat;
        }
        
      .geotest-drop-zone {
        width: 100%;
        min-height: 132px;
        padding: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 28px;
        border: 2px dashed #A8B8EA;
        border-radius: 14px;
        background: #E7F0FF;
    }

    .geotest-drop-icon {
        width: 48px;
        height: 48px;
        object-fit: contain;
        flex-shrink: 0;
    }

    .geotest-drop-text {
        text-align: center;
    }

    .geotest-drop-title {
        color: #253B83;
        font-size: 16px;
        font-weight: 700;
        margin-bottom: 12px;
    }

    .geotest-drop-description {
        color: #64748B;
        font-size: 13px;
    }

      .geotest-dropzone:hover { border-color: #5979E6; background: #EFF3FF; }
      .geotest-table { border: 1px solid #DFE4EF; border-radius: 12px; overflow: hidden; }
      .geotest-table thead tr { background: #F4F6FB; }
      .geotest-table th { font-weight: 700; color: #253B83; }
      .selector-archivos .q-uploader__header { background: #5979E6 !important; box-shadow: none !important; border-radius: 11px !important; }
      .selector-archivos .q-uploader__header:hover { background: #4869D2 !important; }
      
      
    </style>
    ''')
