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

    # taking inputs for constant_fk2d, multiplying_factor, Shear_velocity
    constant_fk2d_value = st.number_input('Enter constant_fk2d_value', step=.01, format="%.2f")

    multiplying_factor_value = st.number_input('Enter multiplying_factor_value', step=.01, format="%.2f")

    Shear_velocity_value = st.number_input('Enter Shear_velocity_value', step=.01, format="%.2f")

    # variables to be used defined
    corr, snr, k_value, lambda_value = 0, 0, 0.0, 0.0

    # button and logic for filtering written
    if st.button("Go to Filter") or st.session_state.filter:
        st.session_state['filter'] = True
        if constant_fk2d_value>0 and multiplying_factor_value>0 and Shear_velocity_value>0:
            st.text(" 1. C\n 2. S\n 3. A\n 4. C & S\n 5. C & A\n 6. S & A\n 7. C & S & A\n 8. All Combine")
            option = st.number_input("Choose filtering method from above", min_value=0, max_value=8, step=1)
            if option == 1:
                corr = st.number_input('Enter Threshold value of C', step=1)
            elif option == 2:
                snr = st.number_input('Enter Threshold value of S', step=1)
            elif option == 3:
                lambda_value = st.number_input('Enter Lambda value of A', step=.01, format="%.2f")
                k_value = st.number_input('Enter k value of A', step=.01, format="%.2f")
            elif option == 4:
                corr = st.number_input('Enter Threshold value of C', step=1)
                snr = st.number_input('Enter Threshold value of S', step=1)
            elif option == 5:
                corr = st.number_input('Enter Threshold value of C', step=1)
                lambda_value = st.number_input('Enter Lambda value of A', step=.01, format="%.2f")
                k_value = st.number_input('Enter k value of A', step=.01, format="%.2f")
            elif option == 6:
                snr = st.number_input('Enter Threshold value of S', step=1)
                lambda_value = st.number_input('Enter Lambda value of A', step=.01, format="%.2f")
                k_value = st.number_input('Enter k value of A', step=.01, format="%.2f")
            elif option == 7 or option == 8:
                corr = st.number_input('Enter Threshold value of C', step=1)
                snr = st.number_input('Enter Threshold value of S', step=1)
                lambda_value = st.number_input('Enter Lambda value of A', step=.01, format="%.2f")
                k_value = st.number_input('Enter k value of A', step=.01, format="%.2f")
            if st.button("Go to Replace") or st.session_state.replace:
                if option>0:
                    st.session_state['replace'] = True
                    st.text(" 1. Previous Point\n 2. 2*last-2nd_last\n 3. Overall_Mean\n 4. 12_Point_Strategy\n 5. Mean Of Previous 2 Points\n 6. All Sequential\n 7. All Parallel\n")
                    replacement_method = st.number_input("Choose Replacement Method From Above ", step=1, min_value=1, max_value=7)
                    if st.button("Compute"):
                        if replacement_method>0:
                            subprocess.run(["python", "psat_v3.py", str(constant_fk2d_value), str(multiplying_factor_value), str(Shear_velocity_value), str(option),str(corr), str(snr), str(lambda_value), str(k_value), str(replacement_method)])
                            st.text("Computed! You can check your Results_v2.csv to see the computed values!!")

take_input()