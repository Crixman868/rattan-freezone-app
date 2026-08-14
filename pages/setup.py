import streamlit as st

st.title("📦 Rattan's Logistics Hub - Device Setup")
st.write("Welcome! Please install the application on your specific device using the instructions below.")

st.divider()

# Windows Section
st.header("🪟 Windows Users")
st.write("1. Click the download button below.")
st.write("2. Open the downloaded file to begin installation.")
st.warning("Note: If Windows Defender pops up, click **'More info'** and then **'Run anyway'** to bypass the one-time warning and install.")
st.link_button("Download for Windows", "https://drive.google.com/file/d/1379YTf25SYsr2phX2smPw1TpEXmZs0-o/view?usp=sharing")

st.divider()

# Android Section
st.header("🤖 Android Users")
st.write("1. Click the download button below to grab the Android package.")
st.write("2. Open the downloaded file to install.")
st.warning("Note: You may need to toggle **'Allow from this source'** in your device settings to successfully sideload the app.")
st.link_button("Download for Android", "https://drive.google.com/file/d/1GBluqmcD3EWk3sIoP22vtMfEyixJcvYN/view?usp=sharing")

st.divider()

# iOS Section
st.header("🍏 iOS (Apple) Users")
st.write("Apple devices connect to the Hub directly through the web without needing a download file. To install:")
st.write("1. Open the main app link in **Safari**.")
st.write("2. Tap the **Share** button at the bottom of the screen (the square with an arrow pointing up).")
st.write("3. Scroll down and select **Add to Home Screen**.")
st.success("This will instantly place the RattanHub app icon on your device's home screen!")