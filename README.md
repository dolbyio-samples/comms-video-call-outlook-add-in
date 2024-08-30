# :warning: This repository is no longer maintained :warning:

# Outlook Add-in For Scheduling Video Conferences

This repository is a ready to use Outlook Add-in sample to schedule instant video calls without having to leave Outlook Calendar. With the click of a button, meetings are scheduled and the relevant meeting information is sent via a calendar invitation. 

You can either run the add-in in your local server and add the manifest.xml file as a custom add-in in your Outlook client or you can deploy it and use the URL of the manifest to upload the add-in to your Outlook app. This repo uses Netlify as an deployment environment, by cloning this project and creating on your own repo, you can host your add-in on Netlify by following [this tutorial](https://www.netlify.com/blog/2016/09/29/a-step-by-step-guide-deploying-on-netlify/).

## Getting Started Steps:
1. Clone this repository. 
2. Make sure you have `npm` installed and run `npm install` in the root directory of the project.
3. Create an .env file to store your meeting provider domain name as DOMAIN_NAME, or you can hardcode the domain name in the `commands.js` file. Having the domain name as an environment variable is useful if you want hide your meeting and change the domain name directly from you deployment environment settings or in an .env file. 
4. If you want to run the add-in from your localhost, run `npm run dev-server` and follow [this tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=windows) for sideloading the manifest.xml to your Outlook client.
5. If you want to deploy your app to your production envioronment, go to `webpack.config.js` line 11, and change it with your own URL.
Then type, `npm run build` and use the dist folder for deployment.
6. To upload the manifest.xml file as an add-in to your Outlook client follow the [Sideload manually guide](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=windows#sideload-manually). To upload the URL of the deployed manifest.xml file, follow the same steps in the [Sideload manually guide](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=windows#sideload-manually), but select `Add from URL` and paste your deployed manifest.xml URL. 

## Deployment Runbook
This repository is hosted on devrel@dolby.com's Netlify account. After applying getting started step 5, and after deploying this repo from Github, netlfiy.toml uses `dist` as publish directory and `npm run build` as build command. If you want to change meeting domain name, you need to change the  environment variable(DOMAIN_NAME) in Netlify site settings. 

<img width="1393" alt="netlify-env" src="https://github.com/dolbyio-samples/comms-video-call-outlook-add-in/assets/63646687/b3199399-bd08-4735-ae37-0d85e798eb9e">

-> To learn more about Outlook add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/ 

-> To learn more about creating add-ins to schedule meetings in Outlook: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/online-meeting?tabs=non-mobile

-> To learn more about creating your own Video Call App with Dolby.io: https://github.com/dolbyio-samples/comms-app-react-videocall

-> Learn how to get started with Dolby.io Communications APIs: https://docs.dolby.io/communications-apis/docshttps://docs.dolby.io/communications-apis/docs
