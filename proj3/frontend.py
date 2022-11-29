import streamlit as st
import subprocess

# sets the title of the page
st.set_page_config(page_title='PSAT_V3')

# sets the session to prevent re-rendering of the whole component.
if 'replace' not in st.session_state:
    st.session_state['replace'] = False

if "filter" not in st.session_state:
    st.session_state['filter'] = False

def take_input():
    constant_fk2d_value = st.number_input('Enter constant_fk2d_value', step=.01, format="%.2f")

    multiplying_factor_value = st.number_input('Enter multiplying_factor_value', step=.01, format="%.2f")

    Shear_velocity_value = st.number_input('Enter Shear_velocity_value', step=.01, format="%.2f")

take_input()