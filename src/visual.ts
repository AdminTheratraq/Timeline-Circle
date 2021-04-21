/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */
"use strict";

import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataViewObjects = powerbi.DataViewObjects;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionIdBuilder = powerbi.extensibility.ISelectionIdBuilder;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import * as d3 from "d3";
import { VisualSettings } from "./settings";
import * as sanitizeHtml from "sanitize-html";
import * as validDataUrl from "valid-data-url";
import { color } from "d3";

export interface TimelineData {
  Title: string;
  Description: string;
  EventStartDate: Date;
  EventEndDate: Date;
  CompanyLink: string;
  HeaderImage: string;
  FooterImage: string;
  selectionId: powerbi.visuals.ISelectionId;
}

export interface Timelines {
  Timeline: TimelineData[];
}

export function logExceptions(): MethodDecorator {
  return (
    target: Object,
    propertyKey: string,
    descriptor: TypedPropertyDescriptor<any>
  ): TypedPropertyDescriptor<any> => {
    return {
      value: function () {
        try {
          return descriptor.value.apply(this, arguments);
        } catch (e) {
          throw e;
        }
      },
    };
  };
}

export function getCategoricalObjectValue<T>(
  objects: DataViewObjects,
  index: number,
  objectName: string,
  propertyName: string,
  defaultValue: T
): T {
  if (objects) {
    let object = objects[objectName];
    if (object) {
      let property: T = <T>object[propertyName];
      if (property !== undefined) {
        return property;
      }
    }
  }
  return defaultValue;
}

export class Visual implements IVisual {
  private target: d3.Selection<HTMLElement, any, any, any>;
  private header: d3.Selection<HTMLElement, any, any, any>;
  private footer: d3.Selection<HTMLElement, any, any, any>;
  private svg: d3.Selection<SVGElement, any, any, any>;
  private margin = { top: 50, right: 40, bottom: 50, left: 40 };
  private settings: VisualSettings;
  private host: IVisualHost;
  private initLoad = false;
  private events: IVisualEventService;
  private xScale: d3.ScaleTime<number, number>;
  private yScale: d3.ScaleLinear<number, number>;
  private gbox: d3.Selection<SVGElement, any, any, any>;
  private colors: any[];
  private selectionManager: ISelectionManager;

  constructor(options: VisualConstructorOptions) {
    this.target = d3.select(options.element);
    this.header = d3.select(options.element).append("div");
    this.svg = d3.select(options.element).append("svg");
    this.footer = d3.select(options.element).append("div");
    this.host = options.host;
    this.events = options.host.eventService;
    this.selectionManager = options.host.createSelectionManager();
  }

  @logExceptions()
  public update(options: VisualUpdateOptions) {
    this.events.renderingStarted(options);
    this.settings = Visual.parseSettings(
      options && options.dataViews && options.dataViews[0]
    );
    this.svg.selectAll("*").remove();
    let _this = this;
    let vpWidth = options.viewport.width;

    this.header.selectAll("img").remove();
    this.header.classed("header", false);
    this.footer.selectAll("img").remove();
    this.footer.classed("footer", false);

    let vpHeight = options.viewport.height;

    if (this.settings.timeline.layout.toLowerCase() === "header" || this.settings.timeline.layout.toLowerCase() === "footer") {
      vpHeight = options.viewport.height - 105;
    }

    this.svg.attr("height", vpHeight);
    this.svg.attr("width", vpWidth);

    let gHeight = vpHeight - this.margin.top - this.margin.bottom;
    let gWidth = vpWidth - this.margin.left - this.margin.right;

    this.target.on("contextmenu", () => {
      const mouseEvent: MouseEvent = <MouseEvent>d3.event;
      const eventTarget: any = mouseEvent.target;
      let dataPoint: any = d3.select(eventTarget).datum();
      this.selectionManager.showContextMenu(
        dataPoint ? dataPoint.selectionId : {},
        {
          x: mouseEvent.clientX,
          y: mouseEvent.clientY,
        }
      );
      mouseEvent.preventDefault();
    });

    let timelineData = Visual.CONVERTER(options.dataViews[0], this.host);
    timelineData = timelineData.slice(0, 100);
    let minDate, maxDate, currentDate;
    let timelineLocalData: TimelineData[] = [];
    currentDate = new Date();

    if (timelineData.length > 0) {
      minDate = new Date(currentDate.getFullYear() - 1, 0, 1);
      timelineLocalData = timelineData.map<TimelineData>((d) => { if (d.EventStartDate.getFullYear() >= minDate.getFullYear()) { return d;} }).filter(e => e);
      maxDate = new Date(currentDate.getFullYear() + 8, 0, 1);
      timelineLocalData = timelineLocalData.map<TimelineData>((d) => { if (d.EventEndDate.getFullYear() <= maxDate.getFullYear()) { return d; } }).filter(e => e);
    }

    if (timelineLocalData.length > 0) {
      timelineData = timelineLocalData;
    } else if (timelineLocalData.length == 0) {
      minDate = new Date(Math.min.apply(null,timelineData.map((d) => d.EventStartDate)));
      maxDate = new Date(Math.max.apply(null, timelineData.map((d) => d.EventEndDate)));
      minDate = new Date(minDate.getFullYear(), 0, 1);
      maxDate = new Date(maxDate.getFullYear() + 1, 0, 1);
    }

    let colors = this.getColors();

    let titleData = timelineData.map((d) => d.Title).filter((v, i, self) => self.indexOf(v) === i);

    let titleColorData = titleData.map((d, i) => {
      if (colors[i]) { return { title: d, color: colors[i]};
      } else {
        let randomNumber: number = this.getRandomNumberBetween(0, 29);
        return { title: d, color: colors[randomNumber] }
      }
    });

    this.renderHeaderAndFooter(timelineData);
    this.renderXandYAxis(minDate, maxDate, gWidth, gHeight);
    this.renderTitle(vpWidth);
    this.defineSVGDefs(titleColorData);
    this.renderXAxisCirclesAndQuarters();

    if (this.settings.timeline.layout.toLowerCase() === "header" || this.settings.timeline.layout.toLowerCase() === "footer") {
      this.renderTimeRangeLines(gHeight, timelineData);
      this.renderGBox(timelineData);
    } else {
      this.renderTimeRangeLinesWithoutLayout(gHeight, timelineData);
      this.renderGBoxWithoutLayout(timelineData);
    }

    this.renderCircles(timelineData, titleColorData);
    this.renderEllipses(titleColorData);
    this.renderText(titleColorData);
    this.handleHyperLinkClick();
    this.renderVisualBorder(vpWidth, vpHeight);
    this.events.renderingFinished(options);
  }

  private renderHeaderAndFooter(timelineData: TimelineData[]) {
    let [timeline] = timelineData;
    if (this.settings.timeline.layout.toLowerCase() === "header") {
      this.header
        .attr("class", "header")
        .append("img")
        .attr(
          "src",
          validDataUrl(timeline.HeaderImage) ? timeline.HeaderImage : ""
        ).exit().remove();
    } else if (this.settings.timeline.layout.toLowerCase() === "footer") {
      this.footer
        .attr("class", "footer")
        .append("img")
        .attr(
          "src",
          validDataUrl(timeline.FooterImage) ? timeline.FooterImage : ""
        );
    }
  }

  private getRandomNumberBetween(min, max) {
    // Create byte array and fill with 1 random number
    var byteArray = new Uint8Array(1);
    window.crypto.getRandomValues(byteArray);

    var range = max - min + 1;
    var max_range = 256;
    if (byteArray[0] >= Math.floor(max_range / range) * range)
      return this.getRandomNumberBetween(min, max);
    return min + (byteArray[0] % range);
  }

  private getColors() {
    return [{ "dark": "#3F5003", "light": "#D0E987", "medium": "#AFD045" }, { "dark": "#252D48", "light": "#81909F", "medium": "#3B4D64" }, { "dark": "#8D4F0F", "light": "#D8A26D", "medium": "#C87825" }, { "dark": "#337779", "light": "#B2DFE0", "medium": "#6FCBCC" }, { "dark": "#003366", "light": "#66ffff", "medium": "#4791AE" }, { "dark": "rgba(49, 27, 146,1)", "light": "rgba(49, 27, 146,0.2)", "medium": "rgba(49, 27, 146,0.5)" }, { "dark": "rgba(245, 127, 23,1)", "light": "rgba(245, 127, 23,0.2)", "medium": "rgba(245, 127, 23,0.5)" }, { "dark": "rgba(183, 28, 28,1)", "light": "rgba(183, 28, 28,0.2)", "medium": "rgba(183, 28, 28,0.5)" }, { "dark": "rgba(136, 14, 79,1)", "light": "rgba(136, 14, 79,0.2)", "medium": "rgba(136, 14, 79,0.5)" }, { "dark": "rgba(27, 94, 32,1)", "light": "rgba(27, 94, 32,0.2)", "medium": "rgba(27, 94, 32,0.5)" }, { "dark": "rgba(255, 0, 0,1)", "light": "rgba(255, 0, 0,0.2)", "medium": "rgba(255, 0, 0,0.5)" }, { "dark": "rgba(0, 0, 255,1)", "light": "rgba(0, 0, 255,0.2)", "medium": "rgba(0, 0, 255,0.5)" }, { "dark": "rgba(0, 255, 0,1)", "light": "rgba(0, 255, 0,0.2)", "medium": "rgba(0, 255, 0,0.5)" }, { "dark": "rgba(94, 89, 27,1)", "light": "rgba(94, 89, 27,0.2)", "medium": "rgba(94, 89, 27,0.5)" }, { "dark": "rgba(27, 94, 91,1)", "light": "rgba(27, 94, 91,0.2)", "medium": "rgba(27, 94, 91,0.5)" }, { "dark": "rgba(11, 101, 153,1)", "light": "rgba(11, 101, 153,0.2)", "medium": "rgba(11, 101, 153,0.5)" }, { "dark": "rgba(11, 45, 153,1)", "light": "rgba(11, 45, 153,0.2)", "medium": "rgba(11, 45, 153,0.5)" }, { "dark": "rgba(114, 11, 153,1)", "light": "rgba(114, 11, 153,0.2)", "medium": "rgba(114, 11, 153,0.5)" }, { "dark": "rgba(153, 11, 134,1)", "light": "rgba(153, 11, 134,0.2)", "medium": "rgba(153, 11, 134,0.5)" }, { "dark": "rgba(249, 5, 134,1)", "light": "rgba(249, 5, 134,0.2)", "medium": "rgba(249, 5, 134,0.5)" }, { "dark": "#800000", "light": "#ff8080", "medium": "#ff3333" }, { "dark": "#808000", "light": "#ffff80", "medium": "#e6e600" }, { "dark": "#a0522d", "light": "#e3b29c", "medium": "#d28460" }, { "dark": "#2f4f4f", "light": "#afcfcf", "medium": "#70a9a9" }, { "dark": "#ff4500", "light": "#ff8f66", "medium": "#ff6933" }, { "dark": "#cf6363", "light": "#e7b1b1", "medium": "#db8a8a" }, { "dark": "#99cd32", "light": "#cce698", "medium": "#aed75b" }, { "dark": "#3f000f", "light": "#ff809d", "medium": "#ff003c" }, { "dark": "#2e8b57", "light": "#9fdfbb", "medium": "#53c685" }, { "dark": "#00cccc", "light": "#99ffff", "medium": "#33ffff" }];
  }

  private renderXandYAxis(minDate, maxDate, gWidth, gHeight) {
    let xAxis;
    this.xScale = d3
      .scaleTime()
      .domain([minDate, maxDate])
      .range([this.margin.left, gWidth]);

    if (this.diff_years(minDate, maxDate) <= 1) {
      xAxis = d3
        .axisBottom(this.xScale)
        .ticks(d3.timeMonth, 1)
        .tickPadding(20)
        .tickFormat(d3.timeFormat("%b'%y"))
        .tickSize(-10);
    } else {
      xAxis = d3
        .axisBottom(this.xScale)
        .ticks(d3.timeYear, 1)
        .tickPadding(20)
        .tickFormat(d3.timeFormat("%Y"))
        .tickSize(-10);
    }

    let xAxisAllTicks = d3
      .axisBottom(this.xScale)
      .ticks(d3.timeMonth, 3)
      .tickPadding(20)
      .tickFormat(d3.timeFormat(""))
      .tickSize(10);

    this.yScale = d3
      .scaleLinear()
      .domain([-100, 100])
      .range([gHeight, this.margin.top]);

    let yAxis = d3.axisLeft(this.yScale);

    let xAxisLineAllTicks = this.svg
      .append("g")
      .attr("class", "x-axis-line-allticks")
      .attr("transform", "translate(" + 20 + "," + (gHeight / 2 + 65) + ")")
      .call(xAxisAllTicks);

    let xAxisLine = this.svg
      .append("g")
      .attr("class", "x-axis-line")
      .attr("transform", "translate(" + 20 + "," + (gHeight / 2 + 65) + ")")
      .call(xAxis);

    this.svg
      .append("g")
      .attr("class", "y-axis")
      .call(yAxis)
      .attr("display", "none");
  }

  private renderTitle(vpWidth) {
    let gTitle = this.svg
      .append("g")
      .attr("x", 0)
      .attr("y", 0)
      .attr("width", vpWidth)
      .attr("height", 35);

    gTitle
      .append("rect")
      .attr("class", "chart-header")
      .attr("width", vpWidth)
      .attr("height", 35);

    gTitle
      .append("text")
      .attr("x", vpWidth / 2)
      .attr("y", 35 / 2)
      .attr("dominant-baseline", "middle")
      .attr("text-anchor", "middle")
      .text(this.settings.timeline.title)
      .attr("fill", "#ffffff")
      .attr("font-size", 24);
  }

  private defineSVGDefs(titleColorData) {
    let svgDefs = this.svg.append("defs");

    titleColorData.forEach((c, i) => {
      if (c.color) {
        let linearGradientTopToBottom = svgDefs
          .append("linearGradient")
          .attr("x2", "0%")
          .attr("y2", "100%")
          .attr("id", "linearGradientTopToBottom" + c.title.replace(/ /g, ""));

        linearGradientTopToBottom
          .append("stop")
          .attr("stop-color", c.color.dark)
          .attr("offset", "0");

        linearGradientTopToBottom
          .append("stop")
          .attr("stop-color", c.color.light)
          .attr("offset", "1");

        let linearGradientBottomToTop = svgDefs
          .append("linearGradient")
          .attr("x2", "0%")
          .attr("y2", "100%")
          .attr("id", "linearGradientBottomToTop" + c.title.replace(/ /g, ""));

        linearGradientBottomToTop
          .append("stop")
          .attr("stop-color", c.color.light)
          .attr("offset", "0");

        linearGradientBottomToTop
          .append("stop")
          .attr("stop-color", c.color.dark)
          .attr("offset", "1");
      }
    });
  }

  private renderXAxisCirclesAndQuarters() {
    let year,
      darkGrey = "#636363",
      lightGrey = "#868686",
      color = "#868686";
    this.svg
      .selectAll(".x-axis-line-allticks .tick")
      .insert("rect")
      .attr("x", 0)
      .attr("y", -25)
      .attr("width", "25%")
      .attr("height", 50)
      .attr("fill", (d: Date, i) => {
        if (i % 4 !== 0) {
          return color;
        } else {
          if (color === lightGrey) {
            color = darkGrey;
          } else {
            color = lightGrey;
          }
          return color;
        }
      });

    this.svg
      .selectAll(".x-axis-line-allticks .tick line")
      .attr("stroke", "#ffffff")
      .attr("stroke-width", 4);

    this.svg
      .selectAll(".x-axis-line .tick")
      .insert("circle")
      .attr("cx", 0)
      .attr("cy", 0)
      .attr("r", 27)
      .attr("stroke", "#525252")
      .attr("stroke-width", 4)
      .attr("fill", "#ffffff");

    this.svg
      .selectAll(".x-axis-line .tick text")
      .attr("y", -5)
      .attr("fill", "#000000")
      .raise();
  }

  private renderTimeRangeLines(gHeight, timelineData) {
    this.svg
      .selectAll(".line")
      .data(timelineData)
      .enter()
      .append("rect")
      .attr("x", (d: any, i) => {
        return this.xScale(d.EventStartDate) + 20;
      })
      .attr("width", "8px")
      .attr("y", (d, i) => {
        if (i % 2 === 0) {
          return this.yScale(-29);
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return this.yScale(49);
          } else {
            return this.yScale(9);
          }
        }
      })
      .attr("height", (d, i) => {
        if (i % 2 === 0) {
          let count = i / 2;
          if (count % 2 === 0) {
            return gHeight - this.yScale(-45);
          } else {
            return gHeight - this.yScale(-85);
          }
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return gHeight - this.yScale(-45);
          } else {
            return gHeight - this.yScale(-85);
          }
        }
      })
      .style("fill", (d: any, i) => {
        if (i % 2 === 0) {
          return (
            "url(#linearGradientTopToBottom" + d.Title.replace(/ /g, "") + ")"
          );
        } else {
          return (
            "url(#linearGradientBottomToTop" + d.Title.replace(/ /g, "") + ")"
          );
        }
      });
    this.svg
      .selectAll(".line")
      .data(timelineData)
      .enter()
      .append("rect")
      .attr("x", (d: any, i) => {
        return this.xScale(d.EventEndDate) + 20;
      })
      .attr("width", "8px")
      .attr("y", (d, i) => {
        if (i % 2 === 0) {
          return this.yScale(-29);
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return this.yScale(49);
          } else {
            return this.yScale(9);
          }
        }
      })
      .attr("height", (d, i) => {
        if (i % 2 === 0) {
          let count = i / 2;
          if (count % 2 === 0) {
            return gHeight - this.yScale(-45);
          } else {
            return gHeight - this.yScale(-85);
          }
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return gHeight - this.yScale(-45);
          } else {
            return gHeight - this.yScale(-85);
          }
        }
      })
      .style("fill", (d: any, i) => {
        if (i % 2 === 0) {
          return (
            "url(#linearGradientTopToBottom" + d.Title.replace(/ /g, "") + ")"
          );
        } else {
          return (
            "url(#linearGradientBottomToTop" + d.Title.replace(/ /g, "") + ")"
          );
        }
      });
  }

  private renderTimeRangeLinesWithoutLayout(gHeight, timelineData) {
    this.svg
      .selectAll(".line")
      .data(timelineData)
      .enter()
      .append("rect")
      .attr("x", (d: any, i) => {
        return this.xScale(d.EventStartDate) + 20;
      }).attr("width", "8px")
      .attr("y", (d, i) => {
        if (i % 2 === 0) {
          return this.yScale(-23);
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return this.yScale(60);
          } else {
            return this.yScale(15);
          }
        }
      }).attr("height", (d, i) => {
        if (i % 2 === 0) {
          let count = i / 2;
          if (count % 2 === 0) {
            return gHeight - this.yScale(-35);
          } else {
            return gHeight - this.yScale(-80);
          }
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return gHeight - this.yScale(-35);
          } else {
            return gHeight - this.yScale(-80);
          }
        }
      })
      .style("fill", (d: any, i) => {
        if (i % 2 === 0) {
          return (
            "url(#linearGradientTopToBottom" + d.Title.replace(/ /g, "") + ")"
          );
        } else {
          return (
            "url(#linearGradientBottomToTop" + d.Title.replace(/ /g, "") + ")"
          );
        }
      });
    this.svg
      .selectAll(".line")
      .data(timelineData)
      .enter()
      .append("rect")
      .attr("x", (d: any, i) => {
        return this.xScale(d.EventEndDate) + 20;
      }).attr("width", "8px")
      .attr("y", (d, i) => {
        if (i % 2 === 0) {
          return this.yScale(-23);
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return this.yScale(60);
          } else {
            return this.yScale(15);
          }
        }
      }).attr("height", (d, i) => {
        if (i % 2 === 0) {
          let count = i / 2;
          if (count % 2 === 0) {
            return gHeight - this.yScale(-35);
          } else {
            return gHeight - this.yScale(-80);
          }
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            return gHeight - this.yScale(-35);
          } else {
            return gHeight - this.yScale(-80);
          }
        }
      }).style("fill", (d: any, i) => {
        if (i % 2 === 0) {
          return (
            "url(#linearGradientTopToBottom" + d.Title.replace(/ /g, "") + ")"
          );
        } else {
          return (
            "url(#linearGradientBottomToTop" + d.Title.replace(/ /g, "") + ")"
          );
        }
      });
  }

  private renderGBox(timelineData) {
    this.gbox = this.svg
      .selectAll(".box")
      .data(timelineData)
      .enter()
      .append("g")
      .attr("fill", "#ffffff")
      .attr("transform", (d: any, i) => {
        let y;
        if (i % 2 === 0) {
          let count = i / 2;
          if (count % 2 === 0) {
            y = this.yScale(-100);
          } else {
            y = this.yScale(-59);
          }
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            y = this.yScale(69);
          } else {
            y = this.yScale(29);
          }
        }
        return (
          "translate(" + (this.xScale(d.EventStartDate) + 25) + " " + y + ")"
        );
      });
  }

  private renderGBoxWithoutLayout(timelineData) {
    this.gbox = this.svg
      .selectAll(".box")
      .data(timelineData)
      .enter()
      .append("g")
      .attr("fill", "#ffffff")
      .attr("transform", (d: any, i) => {
        let y;
        if (i % 2 === 0) {
          let count = i / 2;
          if (count % 2 === 0) {
            y = this.yScale(-97);
          } else {
            y = this.yScale(-53);
          }
        } else {
          let count = Math.ceil(i / 2);
          if (count % 2 === 0) {
            y = this.yScale(69);
          } else {
            y = this.yScale(24);
          }
        }
        return (
          "translate(" + (this.xScale(d.EventStartDate) + 25) + " " + y + ")"
        );
      });
  }

  private renderCircles(timelineData, titleColorData) {
    let _this = this;

    this.gbox
      .selectAll("g")
      .data((d: any, i) => {
        let diff =
          this.xScale(d.EventEndDate) +
          25 -
          (this.xScale(d.EventStartDate) + 25);
        if (
          d.EventStartDate.getTime() === d.EventEndDate.getTime() ||
          diff <= 35
        ) {
          return [d];
        } else {
          return [];
        }
      })
      .enter()
      .append("circle")
      .attr("cx", (d) => {
        let diff =
          this.xScale(d.EventEndDate) +
          25 -
          (this.xScale(d.EventStartDate) + 25);
        if (
          d.EventStartDate.getTime() !== d.EventEndDate.getTime() &&
          diff <= 35
        ) {
          return diff / 2;
        } else {
          return 0;
        }
      })
      .attr("cy", 0)
      .attr("r", 40)
      .attr("stroke", (d: TimelineData) => {
        let companyColor = titleColorData.find((c) => d.Title === c.title);
        return companyColor.color ? companyColor.color.light : "#000000";
      })
      .attr("stroke-width", 2)
      .attr("fill", (d) => {
        return this.settings.timeline.circleBackground === "opaque"
          ? "#ffffff"
          : "rgba(0,0,0,0)";
      });

    this.gbox
      .selectAll("g")
      .data((d: any, i) => {
        let diff =
          this.xScale(d.EventEndDate) +
          25 -
          (this.xScale(d.EventStartDate) + 25);
        if (
          d.EventStartDate.getTime() === d.EventEndDate.getTime() ||
          diff <= 35
        ) {
          return [d];
        } else {
          return [];
        }
      })
      .enter()
      .append("a")
      .append("circle")
      .attr("cx", (d) => {
        let diff =
          this.xScale(d.EventEndDate) +
          25 -
          (this.xScale(d.EventStartDate) + 25);
        if (
          d.EventStartDate.getTime() !== d.EventEndDate.getTime() &&
          diff <= 35
        ) {
          return diff / 2;
        } else {
          return 0;
        }
      })
      .attr("cy", 0)
      .attr("r", 45)
      .attr("stroke", (d: TimelineData) => {
        let companyColor = titleColorData.find((c) => d.Title === c.title);
        return companyColor.color ? companyColor.color.medium : "#000000";
      })
      .attr("stroke-width", 4)
      .attr("fill", (d) => {
        return this.settings.timeline.circleBackground === "opaque"
          ? "#ffffff"
          : "rgba(0,0,0,0)";
      });

    this.gbox.on("mouseenter", function () {
      d3.select(this).raise();
    });

    this.handleCircleOrEllipseClick();
  }

  private handleCircleOrEllipseClick() {
    let _this = this;
    let currentSelection: TimelineData,
      currentElement,
      currentStroke,
      currentStrokeWidth;
    this.gbox.on("click", (d: TimelineData, i, n) => {
      if (currentSelection && currentSelection.selectionId === d.selectionId) {
        _this.selectionManager.clear().then((ids: ISelectionId[]) => {
          let ellipse = d3.select(n[i]).select("ellipse");
          if (!ellipse.empty()) {
            d3.select(n[i])
              .select("ellipse")
              .attr("stroke", currentStroke)
              .attr("stroke-width", currentStrokeWidth);
          } else {
            d3.select(n[i])
              .select("a circle")
              .attr("stroke", currentStroke)
              .attr("stroke-width", currentStrokeWidth);
          }
          currentSelection = undefined;
        });
      } else {
        _this.selectionManager
          .select(d.selectionId)
          .then((ids: ISelectionId[]) => {
            let curelement = d3.select(currentElement).select("ellipse");
            if (!curelement.empty()) {
              d3.select(currentElement)
                .select("ellipse")
                .attr("stroke", currentStroke)
                .attr("stroke-width", currentStrokeWidth);
            } else {
              d3.select(currentElement)
                .select("a circle")
                .attr("stroke", currentStroke)
                .attr("stroke-width", currentStrokeWidth);
            }
            currentSelection = d;
            currentElement = n[i];
            let ellipse = d3.select(n[i]).select("ellipse");
            if (!ellipse.empty()) {
              currentStroke = d3.select(n[i]).select("ellipse").attr("stroke");
              currentStrokeWidth = d3
                .select(n[i])
                .select("ellipse")
                .attr("stroke-width");
              d3.select(n[i])
                .select("ellipse")
                .attr("stroke", currentStroke)
                .attr("stroke-width", 8);
            } else {
              currentStroke = d3.select(n[i]).select("a circle").attr("stroke");
              currentStrokeWidth = d3
                .select(n[i])
                .select("a circle")
                .attr("stroke-width");
              d3.select(n[i])
                .select("a circle")
                .attr("stroke", currentStroke)
                .attr("stroke-width", 8);
            }
          });
      }
    });
  }

  private renderEllipses(titleColorData) {
    this.gbox
      .selectAll("g")
      .data((d: any, i) => {
        let diff =
          this.xScale(d.EventEndDate) +
          25 -
          (this.xScale(d.EventStartDate) + 25);
        if (
          d.EventStartDate.getTime() !== d.EventEndDate.getTime() &&
          diff > 35
        ) {
          return [d];
        } else {
          return [];
        }
      })
      .enter()
      .append("ellipse")
      .attr("cx", (d: TimelineData, i) => {
        let diff =
          this.xScale(d.EventEndDate) +
          25 -
          (this.xScale(d.EventStartDate) + 25);
        return diff / 2;
      })
      .attr("cy", 2)
      .attr("rx", (d: TimelineData, i) => {
        return (
          this.xScale(d.EventEndDate) +
          25 -
          (this.xScale(d.EventStartDate) + 25)
        );
      })
      .attr("ry", 50)
      .attr("stroke", (d: TimelineData) => {
        let companyColor = titleColorData.find((c) => d.Title === c.title);
        return companyColor ? companyColor.color.light : "#000000";
      })
      .attr("stroke-width", 2)
      .attr("fill", (d) => {
        return this.settings.timeline.circleBackground === "opaque"
          ? "#ffffff"
          : "rgba(0,0,0,0)";
      });
  }

  private renderText(titleColorData) {
    // we are not able to use append method instead of html method here since the Description field has the value as HTML (rich text)
    this.gbox
      .append("foreignObject")
      .html((d: TimelineData) => {
        let companyColor = titleColorData.find((c) => d.Title === c.title);
        let color = companyColor.color ? companyColor.color.medium : "#000000";
        let company =
          '<div><a href="' +
          sanitizeHtml(d.CompanyLink) +
          '" style="color:' +
          color +
          ';">' +
          (d.Title ? sanitizeHtml(d.Title) : "") +
          "</a></div>";
        return (
          '<div title="' +
          this.extractContent(sanitizeHtml(d.Description)) +
          '">' +
          company +
          sanitizeHtml(d.Description) +
          "</div>"
        );
      })
      .attr("x", (d: TimelineData) => {
        if (d.EventStartDate.getTime() === d.EventEndDate.getTime()) {
          return -35;
        } else {
          return -20;
        }
      })
      .attr("y", "-50")
      .attr("width", (d: TimelineData) => {
        if (
          d.EventStartDate.getTime() === d.EventEndDate.getTime() ||
          this.diff_years(d.EventEndDate, d.EventStartDate) < 1
        ) {
          return 70;
        } else {
          let diff =
            this.xScale(d.EventEndDate) +
            25 -
            (this.xScale(d.EventStartDate) + 25);
          return diff + diff / 2;
        }
      })
      .attr("height", 60)
      .attr("fill", "#000000")
      .attr("transform", "translate(0,20)")
      .attr("font-size", 10)
      .attr("font-weight", "bold");
  }

  private handleHyperLinkClick() {
    let _this = this;
    let baseurl = "https://strategicanalysisinc.sharepoint.com";
    this.svg.selectAll("foreignObject a").on("click", function (e: Event) {
      e = e || window.event;
      let target: any = e.target || e.srcElement;
      let link = d3.select(this).attr("href");
      if (link.indexOf("http") === -1 || link.indexOf("http") > 0) {
        link = baseurl + link;
      }
      _this.host.launchUrl(link);
      d3.event.preventDefault();
      return false;
    });
  }

  private renderVisualBorder(vpWidth, vpHeight) {
    this.svg
      .append("rect")
      .attr("class", "visual-border-rect")
      .attr("x", 0)
      .attr("y", 0)
      //.attr('transform', 'translate(' + (this.margin.left - 29) + ',' + (this.margin.top - 35) + ')')
      .attr("width", vpWidth)
      .attr("height", vpHeight)
      .attr("stroke-width", "2px")
      .attr("stroke", "#333")
      .attr("fill", "transparent");
  }

  private extractContent(str) {
    if (str === null || str === "") return false;
    else str = str.toString();
    return str.replace(/(<([^>]+)>)/gi, "");
  }

  // converter to table data
  public static CONVERTER(
    dataView: DataView,
    host: IVisualHost
  ): TimelineData[] {
    let resultData: TimelineData[] = [];
    let tableView = dataView.table;
    let _rows = tableView.rows;
    let _columns = tableView.columns;
    let _titleIndex = -1,
      _descIndex = -1,
      _startDateIndex = -1,
      _endDateIndex = -1,
      _companyLinkIndex = -1,
      _headerImageIndex = -1,
      _footerImageIndex = -1;
    for (let ti = 0; ti < _columns.length; ti++) {
      if (_columns[ti].roles.hasOwnProperty("Title")) {
        _titleIndex = ti
      } else if (_columns[ti].roles.hasOwnProperty("Description")) {
        _descIndex = ti;
      } else if (_columns[ti].roles.hasOwnProperty("EventStartDate")) {
        _startDateIndex = ti;
      } else if (_columns[ti].roles.hasOwnProperty("EventEndDate")) {
        _endDateIndex = ti;
      } else if (_columns[ti].roles.hasOwnProperty("CompanyLink")) {
        _companyLinkIndex = ti;
      } else if (_columns[ti].roles.hasOwnProperty("HeaderImage")) {
        _headerImageIndex = ti;
      } else if (_columns[ti].roles.hasOwnProperty("FooterImage")) {
        _footerImageIndex = ti;
      }
    }
    for (let i = 0; i < _rows.length; i++) {
      let row = _rows[i];
      let dp = {
        Title: row[_titleIndex] ? row[_titleIndex].toString() : null,
        Description: row[_descIndex] ? row[_descIndex].toString() : null,
        EventStartDate: row[_startDateIndex]
          ? new Date(Date.parse(row[_startDateIndex].toString()))
          : null,
        EventEndDate: row[_endDateIndex]
          ? new Date(Date.parse(row[_endDateIndex].toString()))
          : null,
        CompanyLink: row[_companyLinkIndex]
          ? row[_companyLinkIndex].toString()
          : null,
        HeaderImage: row[_headerImageIndex]
          ? row[_headerImageIndex].toString()
          : null,
        FooterImage: row[_footerImageIndex]
          ? row[_footerImageIndex].toString()
          : null,
        selectionId: host
          .createSelectionIdBuilder()
          .withTable(tableView, i)
          .createSelectionId(),
      };
      resultData.push(dp);
    }

    return resultData;
  }

  private static parseSettings(dataView: DataView): VisualSettings {
    return <VisualSettings>VisualSettings.parse(dataView);
  }

  /**
   * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
   * objects and properties you want to expose to the users in the property pane.
   *
   */
  public enumerateObjectInstances(
    options: EnumerateVisualObjectInstancesOptions
  ): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
    return VisualSettings.enumerateObjectInstances(
      this.settings || VisualSettings.getDefault(),
      options
    );
  }

  private diff_years(dt2, dt1) {
    let diff = (dt2.getTime() - dt1.getTime()) / 1000;
    diff /= 60 * 60 * 24;
    return Math.abs(Math.round(diff / 365.25));
  }
}
