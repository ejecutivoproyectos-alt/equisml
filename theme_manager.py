import streamlit as st
import colorsys


def hsla_to_hex(h, s, l, a=1):
    h = max(0, min(360, h)) / 360
    s = max(0, min(100, s)) / 100
    l = max(0, min(100, l)) / 100

    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "{:02X}{:02X}{:02X}".format(int(r * 255), int(g * 255), int(b * 255))


def get_theme_palettes():
    return {
        "Rojo clásico": {
            "primary": (0, 100, 50, 1),
            "secondary": (22, 37, 83, 1),
            "border": (0, 0, 0, 1),
            "title": (0, 100, 50, 1),
            "text_light": (0, 0, 100, 1),
            "text_dark": (0, 0, 0, 1),
            "neutral": (0, 0, 100, 1),
        },
        "Azul corporativo": {
            "primary": (210, 100, 36, 1),
            "secondary": (210, 35, 86, 1),
            "border": (210, 25, 22, 1),
            "title": (210, 100, 36, 1),
            "text_light": (0, 0, 100, 1),
            "text_dark": (0, 0, 0, 1),
            "neutral": (0, 0, 100, 1),
        },
        "Verde institucional": {
            "primary": (145, 65, 34, 1),
            "secondary": (145, 30, 85, 1),
            "border": (145, 20, 20, 1),
            "title": (145, 65, 34, 1),
            "text_light": (0, 0, 100, 1),
            "text_dark": (0, 0, 0, 1),
            "neutral": (0, 0, 100, 1),
        },
        "Dorado sobrio": {
            "primary": (43, 100, 47, 1),
            "secondary": (43, 55, 88, 1),
            "border": (35, 40, 28, 1),
            "title": (43, 100, 37, 1),
            "text_light": (0, 0, 100, 1),
            "text_dark": (0, 0, 0, 1),
            "neutral": (0, 0, 100, 1),
        },
    }


def build_theme_from_palette_name(palette_name):
    palettes = get_theme_palettes()

    if not palette_name or palette_name not in palettes:
        palette_name = "Rojo clásico"

    raw = palettes[palette_name]

    return {
        "name": palette_name,
        "primary": hsla_to_hex(*raw["primary"]),
        "secondary": hsla_to_hex(*raw["secondary"]),
        "border": hsla_to_hex(*raw["border"]),
        "title": hsla_to_hex(*raw["title"]),
        "text_light": hsla_to_hex(*raw["text_light"]),
        "text_dark": hsla_to_hex(*raw["text_dark"]),
        "neutral": hsla_to_hex(*raw["neutral"]),
    }


def init_theme_session():
    if "selected_palette_name" not in st.session_state:
        st.session_state["selected_palette_name"] = "Rojo clásico"

    if "theme" not in st.session_state:
        st.session_state["theme"] = build_theme_from_palette_name(
            st.session_state["selected_palette_name"]
        )


def render_global_theme_selector():
    init_theme_session()

    palettes = list(get_theme_palettes().keys())

    st.sidebar.markdown("## Diseño del documento")

    selected = st.sidebar.selectbox(
        "Paleta global",
        palettes,
        index=palettes.index(st.session_state["selected_palette_name"]),
        key="global_palette_select"
    )

    st.session_state["selected_palette_name"] = selected
    st.session_state["theme"] = build_theme_from_palette_name(selected)

    theme = st.session_state["theme"]

    st.sidebar.markdown(
        f"""
        <div style="display:flex; gap:8px; margin-top:8px; margin-bottom:8px;">
            <div style="width:28px; height:28px; background:#{theme['primary']}; border:1px solid #000;"></div>
            <div style="width:28px; height:28px; background:#{theme['secondary']}; border:1px solid #000;"></div>
            <div style="width:28px; height:28px; background:#{theme['border']}; border:1px solid #000;"></div>
            <div style="width:28px; height:28px; background:#{theme['neutral']}; border:1px solid #000;"></div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.sidebar.caption(f"Paleta activa: {theme['name']}")


def get_active_theme():
    init_theme_session()
    return st.session_state["theme"]