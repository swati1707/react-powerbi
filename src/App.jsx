import React from "react";
import { service, factories, models, IEmbedConfiguration } from "powerbi-client";
import * as config from "./config";

const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

let accessToken = "";
let embedUrl = "";
let embedToken = "";
let reportContainer;
let reportRef;
let loading;

class App extends React.Component {

    constructor(props) {
        super(props);

        this.state = { accessToken: "", embedUrl: "", error: [], embedToken: "" };

        reportRef = React.createRef();

        // Report container
        loading = (
            <div
                id="reportContainer"
                ref={reportRef} >
                Loading the report...
            </div>
        );
    }

    componentDidMount() {

        if (reportRef !== null) {
            reportContainer = reportRef["current"];
        }

        if (config.workspaceId === "" || config.reportId === "") {
            this.setState({ error: ["Please assign values for workspace id and report id"] })
        } else {
            this.getAccessToken();

        }
    }

    getAccessToken() {
        const thisObj = this;

        fetch(" https://login.microsoftonline.com/common/oauth2/v2.0/token", {
            headers: {
                "Content-type": "application/x-www-form-urlencoded"
            },
            method: "GET",
            body: {
                grant_type: "client_credentials",
                client_id: config.clientId,
                client_secret: config.clientSecret,
                scope: "https://analysis.windows.net/powerbi/api/.default"

            }
        })
            .then(function (response) {
                const errorMessage = [];
                errorMessage.push("Error occurred while fetching the access token of the report")
                errorMessage.push("Request Id: " + response.headers.get("requestId"));

                response.json()
                    .then(function (body) {
                        // Successful response
                        if (response.ok) {
                            accessToken = body["accessToken"];
                            thisObj.setState({ accessToken: accessToken });
                        }
                        // If error message is available
                        else {
                            errorMessage.push("Error " + response.status + ": " + body.error.code);

                            thisObj.setState({ error: errorMessage });
                        }

                    })
                    .catch(function () {
                        errorMessage.push("Error " + response.status + ":  An error has occurred");

                        thisObj.setState({ error: errorMessage });
                    });
            })
            .catch(function (error) {

                // Error in making the API call
                thisObj.setState({ error: error });
            })
    }

    getEmbedUrl() {
        const thisObj = this;

        fetch("https://api.powerbi.com/v1.0/myorg/groups/" + config.workspaceId + "/reports/" + config.reportId, {
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            method: "GET"
        })
            .then(function (response) {
                const errorMessage = [];
                errorMessage.push("Error occurred while fetching the embed URL of the report")
                errorMessage.push("Request Id: " + response.headers.get("requestId"));

                response.json()
                    .then(function (body) {
                        // Successful response
                        if (response.ok) {
                            embedUrl = body["embedUrl"];
                            thisObj.setState({ accessToken: accessToken, embedUrl: embedUrl });
                        }
                        // If error message is available
                        else {
                            errorMessage.push("Error " + response.status + ": " + body.error.code);

                            thisObj.setState({ error: errorMessage });
                        }

                    })
                    .catch(function () {
                        errorMessage.push("Error " + response.status + ":  An error has occurred");

                        thisObj.setState({ error: errorMessage });
                    });
            })
            .catch(function (error) {

                // Error in making the API call
                thisObj.setState({ error: error });
            })
    }

    getEmbedToken() {
      const thisObj = this;

      fetch("https://api.powerbi.com/v1.0/myorg/GenerateToken" , {
          headers: {
              "Authorization": "Bearer " + this.state.accessToken
          },
          method: "POST",
          body: {
            "datasets": [
              {
                "id": config.datasetId
              }
            ],
            "reports": [
              {
                "id": config.reportId
              }
            ]
          }
      })
          .then(function (response) {
              const errorMessage = [];
              errorMessage.push("Error occurred while fetching the embed token of the report")
              errorMessage.push("Request Id: " + response.headers.get("requestId"));

              response.json()
                  .then(function (body) {
                      // Successful response
                      if (response.ok) {
                          embedToken = body["embedToken"];
                          thisObj.setState({ accessToken: accessToken, embedToken: embedToken });
                      }
                      // If error message is available
                      else {
                          errorMessage.push("Error " + response.status + ": " + body.error.code);

                          thisObj.setState({ error: errorMessage });
                      }

                  })
                  .catch(function () {
                      errorMessage.push("Error " + response.status + ":  An error has occurred");

                      thisObj.setState({ error: errorMessage });
                  });
          })
          .catch(function (error) {

              // Error in making the API call
              thisObj.setState({ error: error });
          })
    }

    render() {

        if(this.state.accessToken !== "") {
            this.getEmbedToken();
            this.getEmbedUrl();
        }

        if (this.state.error.length) {
            reportContainer.textContent = "";
            this.state.error.forEach(line => {
                reportContainer.appendChild(document.createTextNode(line));
                reportContainer.appendChild(document.createElement("br"));
            });
        }
        else if (this.state.accessToken !== "" && this.state.embedUrl !== "" && this.state.embedToken !== "") {

            const countryFilter = {
                $schema: "http://powerbi.com/product/schema#basic",
                target: {
                  table: "Country",
                  column: "country_name"
                },
                operator: "In",
                values: [INDIA],
                filterType: 1,
                requireSingleSelection: false
            }
            const embedConfiguration = {
                type: "report",
                tokenType: models.TokenType.Embed,
                accessToken,
                embedUrl,
                id: config.reportId,
                settings: {
                    background: models.BackgroundType.Transparent
                }
                
            };

            const report = powerbi.embed(reportContainer, embedConfiguration);

            report.off("loaded");

            report.on("loaded", function () {
               const existingFilters = report.getFilters()
                .then(function(response) {
                    const newFiltersArr = [...response, countryFilter];
                    report.setFilters(newFiltersArr);
                });
            });

            report.off("rendered");

            report.on("rendered", function () {
                console.log("Report render successful");
            });

            report.off("error");

            report.on("error", function (event) {
                const errorMsg = event.detail;
                console.error(errorMsg);
            });
        }

        return loading;
    }

    componentWillUnmount() {
        powerbi.reset(reportContainer);
    }
}

export default App;