import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
from win32com.client import Dispatch
import pandas as pd
import pythoncom  # Added pythoncom for COM initialization


def speak(text):
    pythoncom.CoInitialize()  # Initialize COM
    try:
        speaker = Dispatch("SAPI.SpVoice")
        speaker.Speak(text)
    finally:
        pythoncom.CoUninitialize()  # Clean up COM


# Load the model and vectorizer
model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

def main():
    st.title(":green[SMS Spam Detection]")
    
    # Sidebar options
    activities = ["Classification", "About"]
    choices = st.sidebar.selectbox("Select Activities", activities)
    
    # Classification functionality
    if choices == "Classification":
        st.subheader(":violet[Classification]")
        msg = st.text_input("Enter a text")
        if st.button("Process"):
            data = [msg]
            vec = cv.transform(data).toarray()
            result = model.predict(vec)
            
            # Check if it's spam or not
            if result[0] == 0:
                st.success("This is Not A Spam SMS", icon="✅")
                speak("This is Not A Spam ")
                st.metric(label="HAM", value="95% SAFE", delta="GOOD")
                
            else:
                st.error("This is A Spam SMS", icon="🚨")
                speak("This is A Spam SMS")
                st.metric(label="SPAM", value="-25% SAFE", delta="-BAD")
    
   
        
# Run the main function
if __name__ == "__main__":
    main()
