import Toc, { TocItem } from "../toc";
import * as React from "react";
import manageJss, { ComponentStyles, DesignSystemProvider, IJSSManagerProps, IManagedClasses } from "@microsoft/fast-jss-manager-react";
import { glyphBuildingblocks, glyphGlobalnavbutton, glyphTransparency } from "@microsoft/fast-glyphs-msft";
import Form from "@microsoft/fast-form-generator-react";
import { get, has, uniqueId } from "lodash-es";
import devSiteDesignSystemDefaults, { IDevSiteDesignSystem } from "../design-system";
import Shell, { ShellHeader, ShellInfoBar, ShellPaneCollapse, ShellSlot } from "../shell";
import { BrowserRouter, Redirect, Route, Switch } from "react-router-dom";
import { ellipsis, localizeSpacing, toPx } from "@microsoft/fast-jss-utilities";
import ComponentWrapper from "./component-wrapper";
import CategoryList from "./category-list";
import SiteTitle from "./title";
import SiteTitleBrand from "./title-brand";
import SiteMenu from "./menu";
import SiteMenuItem from "./menu-item";
import SiteCategory, { Status } from "./category";
import SiteCategoryDocumentation from "./category-documentation";
import SiteCategoryIcon from "./category-icon";
import SiteCategoryItem from "./category-item";
import ActionBar from "./action-bar";
import DevTools, { Framework } from "./dev-tools";
import ConfigurationPanel from "./configuration-panel";
import NotFound from "./not-found";
import ComponentView, { ComponentViewTypes } from "./component-view";
import {
    Canvas,
    Container,
    IContainerClassNamesContract,
    IPaneClassNamesContract,
    Pane,
    PaneResizeDirection,
    Row,
    RowResizeDirection
} from "@microsoft/fast-layouts-react";
import { Direction, isRTL } from "@microsoft/fast-application-utilities";

export enum ComponentViewSlot {
    example = "canvas-example-view",
    detailExample = "canvas-detail-view-example",
    detailDocumentation = "canvas-detail-view-documentation"
}

export interface ISiteProps {
    formChildOptions: IFormChildOption[];
    onUpdateDirection?: (ltr: Direction) => void;
    onUpdateTheme?: (theme: string) => void;
    locales?: string[];
    themes?: ITheme[];
    activeTheme?: ITheme;
    frameworks?: Framework | Framework[];
    activeFramework?: Framework;
    collapsed?: boolean;
    componentBackgroundTransparent?: boolean;
    showComponentStatus?: boolean;
    showTransparencyToggle?: boolean;
}

export interface IFormChildOption {
    component: any;
    schema: any;
    name: string;
}

export interface IComponentRoute {
    route: string;
    schema: any;
    componentMapping: any;
    exampleView: JSX.Element[];
    detailView: JSX.Element[];
    status: Status;
}

export interface IComponentData {
    [T: string]: any[];
}

export interface ISiteState {
    currentPath: string;
    activeComponentIndex: number;
    componentName: string;
    componentData: IComponentData;
    componentStatus: Status;
    detailViewComponentData: IComponentData;
    tableOfContentsCollapsed: boolean;
    componentView: ComponentViewTypes;
    componentBackgroundTransparent: boolean;
    formView: boolean;
    devToolsView: boolean;
    locale: string;
    theme: ITheme;
}

export enum SiteSlot {
    category = "category",
    categoryIcon = "category-icon"
}

export interface ITheme {
    id: string;
    displayName: string;
    background?: string;
}

export interface ISiteManagedClasses {
    "@global": string;
    site_canvasContent: string;
    site_headerTitle: string;
    site_infoBarConfiguration: string;
    site_infoBarConfiguration_base: string;
    site_infoBarConfiguration_input: string;
    site_infoBarConfiguration_theme: string;
    site_paneForm: string;
    site_paneToc: string;
    site_paneTocRow: string;
    site_paneTocTitle: string;
    site_paneToggleButton: string;
    site_paneToggleButtonIcon: string;
    site_paneToggleButtonIconLayout: string;
    site_transparencyToggleButton: string;
    site_transparencyToggleButtonIcon: string;
    site_statusBar: string;
    site_statusComponentName: string;
    site_statusIndicator: string;
    site_statusReleased: string;
    site_statusAlpha: string;
    site_statusBeta: string;
}

const styles: ComponentStyles<ISiteManagedClasses, IDevSiteDesignSystem> = {
    "@global": {
        "body, html": {
            fontFamily: "Segoe UI, SegoeUI, Helvetica Neue, Helvetica, Arial, sans-serif",
            fontSize: toPx(14),
            padding: toPx(0),
            margin: toPx(0)
        }
    },
    site_canvasContent: {
        height: `calc(100% - ${toPx(40)})`,
        display: "flex",
        flexDirection: "column"
    },
    site_headerTitle: {
        fontSize: toPx(15),
        marginLeft: toPx(8),
        cursor: "default"
    },
    site_infoBarConfiguration: {
        display: "flex",
        alignItems: "center",
        padding: toPx(4),
    },
    site_infoBarConfiguration_base: {
        position: "relative",
        "&::before, &::after": {
            content: "''",
            position: "absolute",
            top: toPx(11),
            zIndex: "1",
            borderRadius: toPx(2),
            width: toPx(1),
            height: toPx(10),
            background: "#000000"
        },
        "&::before": {
            right: toPx(15),
            transform: "rotate(45deg)"
        },
        "&::after": {
            right: toPx(22),
            transform: "rotate(-45deg)"
        }
    },
    site_infoBarConfiguration_theme: {
        marginRight: toPx(4)
    },
    site_infoBarConfiguration_input: {
        lineHeight: toPx(16),
        fontSize: toPx(14),
        backgroundColor: "rgba(0, 0, 0, 0.04)",
        borderRadius: toPx(2),
        boxShadow: `inset 0 0 ${toPx(4)} 0 rgba(0, 0, 0, 0.08)`,
        appearance: "none",
        padding: localizeSpacing(Direction.ltr)(`${toPx(8)} ${toPx(36)} ${toPx(8)} ${toPx(10)}`),
        border: "none",
        outline: "none",
        "&:-ms-expand": {
            display: "none"
        },
        "&:hover": {
            boxShadow: `inset 0 0 ${toPx(2)} 0 rgba(0,0,0, .3)`
        },
        "&:focus": {
            boxShadow: (config: IDevSiteDesignSystem): string => {
                return `inset 0 0 0 1 ${config.brandColor}`;
            }
        }
    },
    site_paneForm: {
        padding: toPx(12)
    },
    site_paneToc: {
        padding: "0"
    },
    site_paneTocRow: {
        display: "flex",
        flexWrap: "nowrap",
        alignItems: "center",
    },
    site_paneTocTitle: {
        fontWeight: "bold",
        marginLeft: toPx(-8),
        textOverflow: ellipsis().textOverflow,
        whiteSpace: "nowrap",
        overflow: "hidden",
        cursor: "default"
    },
    site_paneToggleButton: {
        border: "none",
        background: "none",
        padding: toPx(12),
        outline: "0"
    },
    site_paneToggleButtonIcon: {
        height: toPx(16),
        width: toPx(16),
        justifyContent: "center",
        fontSize: toPx(16),
        paddingTop: toPx(2),
        display: "inline-block"
    },
    site_paneToggleButtonIconLayout: {
        height: toPx(40),
        width: toPx(40),
        display: "flex",
        alignItems: "center",
        justifyContent: "center"
    },
    site_transparencyToggleButton: {
        border: "none",
        borderRadius: toPx(2),
        background: "none",
        height: toPx(32),
        width: toPx(32),
        cursor: "pointer",
        outline: "0",
        marginRight: toPx(4),
        opacity: ".85",
        paddingTop: toPx(4),
        "&:hover": {
            opacity: "1"
        },
        "&[aria-pressed=\"true\"]": {
            background: "#FFFFFF",
            opacity: "1"
        }
    },
    site_transparencyToggleButtonIcon: {
        position: "relative",
        top: toPx(1)
    },
    site_statusBar: {
        display: "flex",
        alignItems: "center",
        lineHeight: "1",
        textAlign: "left",
        flex: "1"
    },
    site_statusComponentName: {
        marginLeft: toPx(12),
        marginRight: toPx(15),
    },
    site_statusIndicator: {
        borderRadius: "50%",
        height: toPx(16),
        width: toPx(16),
        marginRight: toPx(4),
        boxSizing: "border-box",
        border: `${toPx(1)} solid #FFFFFF`,
    },
    site_statusReleased: {
        backgroundColor: "#3EC28F",
    },
    site_statusAlpha: {
        backgroundColor: "#824ED2",
    },
    site_statusBeta: {
        backgroundColor: "#950811",
    }
};

class Site extends React.Component<ISiteProps & IManagedClasses<ISiteManagedClasses>, ISiteState> {

    public static defaultProps: Partial<ISiteProps> = {
        locales: ["en", "en-rtl"],
        showComponentStatus: true,
        showTransparencyToggle: true
    };

    public static getDerivedStateFromProps(props: ISiteProps, state: ISiteState): Partial<ISiteState> | null {
        const updatedTheme: ITheme | null = props.activeTheme
            ? props.activeTheme
            : props.themes
                ? props.themes.find((theme: ITheme) => theme.id === state.theme.id)
                : null;

        return updatedTheme === null ? null : { theme: updatedTheme };
    }

    constructor(props: ISiteProps & IManagedClasses<ISiteManagedClasses>) {
        super(props);

        this.state = {
            currentPath: this.initialPath,
            activeComponentIndex: 0,
            tableOfContentsCollapsed: this.props.collapsed || false,
            componentBackgroundTransparent: this.props.componentBackgroundTransparent || false,
            componentView: ComponentViewTypes.examples,
            componentName: this.getComponentName(this.initialPath),
            componentData: this.getComponentData(),
            componentStatus: this.getComponentStatus(this.initialPath),
            detailViewComponentData: this.getDetailViewComponentData(),
            formView: true,
            devToolsView: false,
            locale: "en",
            theme: this.props.activeTheme || this.getInitialTheme()
        };
    }
    public renderNewFormat(): JSX.Element {
        const currentRoute: IComponentRoute = this.routeByPath(this.state.currentPath);

        if (typeof currentRoute !== "object" || currentRoute === null) {
            return null
        }

        const paneStyleSheet: Partial<ComponentStyles<IPaneClassNamesContract, IDevSiteDesignSystem>> = {
            pane: {
                backgroundColor: (config: IDevSiteDesignSystem): string => {
                    return config.lightGray;
                }
            }
        };

        return (
            <Row fill={true}>
                <Pane
                    collapsed={this.state.tableOfContentsCollapsed}
                    resizable={true}
                    resizeFrom={PaneResizeDirection.east}
                    jssStyleSheet={paneStyleSheet}
                    minWidth={200}
                >
                    <div>
                        <button
                            onClick={this.handlePaneCollapse}
                            className={this.props.managedClasses.site_paneToggleButton}
                        >
                            <span
                                className={this.props.managedClasses.site_paneToggleButtonIcon}
                                dangerouslySetInnerHTML={{ __html: glyphGlobalnavbutton }}
                            />

                        </button>
                        {this.renderChildrenBySlot(this, ShellSlot.pane)}
                        <ul className={this.props.managedClasses.site_paneToc}>
                            {this.renderToc()}
                        </ul>
                    </div>
                </Pane>
                <Canvas>
                    <Row>
                        <ActionBar
                            onComponentViewChange={this.onComponentViewChange}
                            onFormToggle={this.onFormToggle}
                            onDevToolsToggle={this.onDevToolsToggle}
                            componentView={this.state.componentView}
                            formView={this.state.formView}
                            devToolsView={this.state.devToolsView}
                        />
                    </Row>
                    <div className={this.props.managedClasses.site_canvasContent}>
                        <ComponentView viewType={this.state.componentView}>
                            {this.renderChildrenBySlot(this, ShellSlot.canvas)}
                            {this.renderComponentByRoute(currentRoute)}
                        </ComponentView>
                        <Row
                            resizable={true}
                            hidden={!this.state.devToolsView}
                            resizeFrom={RowResizeDirection.north}
                            minHeight={180}
                        >
                            {this.renderDevTools(currentRoute.schema)}
                        </Row>
                    </div>
                </Canvas>
                <Pane
                    hidden={!this.state.formView}
                    resizable={true}
                    resizeFrom={PaneResizeDirection.west}
                    jssStyleSheet={paneStyleSheet}
                    minWidth={324}
                >
                    {this.generateForm(currentRoute.schema)}
                </Pane>
            </Row>

        );
    }

    public render(): JSX.Element {
        return (
            <DesignSystemProvider designSystem={devSiteDesignSystemDefaults}>
                <BrowserRouter>
                    <Shell>
                        <ShellHeader>
                            {this.renderChildrenBySlot(this, ShellSlot.header)}
                            {this.renderSiteTitle()}
                        </ShellHeader>
                        {this.renderNewFormat()}
                        { /* TODO: delete Switch */ }
                        <Switch>
                            <Route
                                exact={true}
                                path={"/"}
                            >
                                <Redirect to={this.initialPath} />
                            </Route>
                            <Route path="*" component={NotFound} />
                        </Switch>
                        <ShellInfoBar>
                            <div className={this.props.managedClasses.site_statusBar}>
                                <span className={this.props.managedClasses.site_statusComponentName}>
                                    Component: {this.state.componentName}
                                </span>
                                {this.renderStatus()}
                            </div>

                            {this.renderChildrenBySlot(this, ShellSlot.infoBar)}

                            <div className={this.props.managedClasses.site_infoBarConfiguration}>
                                {this.renderTransparencyToggle()}
                                {this.renderThemeSelect()}
                                <span className={this.props.managedClasses.site_infoBarConfiguration_base}>
                                    <select
                                        className={this.props.managedClasses.site_infoBarConfiguration_input}
                                        onChange={this.handleLocaleUpdate}
                                        value={this.state.locale}
                                    >
                                        {this.renderLocales()}
                                    </select>
                                </span>
                            </div>
                        </ShellInfoBar>
                    </Shell>
                </BrowserRouter>
            </DesignSystemProvider>
        );
    }

    public componentDidMount(): void {
        // If the path we load the site in doesn't match component view, update state
        // to match the path
        if (this.getComponentViewTypesByLocation() !== this.state.componentView) {
            this.setState({
                componentView: this.getComponentViewTypesByLocation()
            });
        }
    }

    public componentDidUpdate(prevProps: ISiteProps): void {
        if (prevProps !== this.props) {
            this.setState({
                componentData: this.getComponentData(),
                detailViewComponentData: this.getDetailViewComponentData()
            });
        }
    }

    private getInitialTheme(): ITheme {
        if (Array.isArray(this.props.themes)) {
            return this.props.themes[0];
        }
    }

    /**
     * Gets the component data for each of the routes
     */
    private getComponentData(): IComponentData {
        const componentData: IComponentData = {};

        this.getRoutes(this.props.children, "/", SiteSlot.category).forEach((route: IComponentRoute) => {
            componentData[route.route] = [];
            route.exampleView.forEach((routeChild: JSX.Element, index: number) => {
                componentData[route.route][index] = routeChild.props.data;
            });
        });

        return componentData;
    }

    private getComponentStatus(currentPath?: string): Status {
        let componentStatus: Status = Status.beta;

        this.getRoutes(this.props.children, "/", SiteSlot.category).forEach((route: IComponentRoute) => {
            if ((currentPath && route.route === currentPath || !currentPath && route.route === this.state.currentPath) && route.status) {
                componentStatus = route.status;
            }
        });

        return componentStatus;
    }

    private getDetailViewComponentData(): IComponentData {
        const componentData: IComponentData = {};

        this.getRoutes(this.props.children, "/", SiteSlot.category).forEach((route: IComponentRoute) => {
            componentData[route.route] = [];
            route.detailView.forEach((routeChild: JSX.Element, index: number) => {
                if (routeChild && routeChild.props && routeChild.props.slot === ComponentViewSlot.detailExample) {
                    componentData[route.route] = routeChild.props.data;
                }
            });
        });

        return componentData;
    }

    private getComponentName = (currentPath?: string): string => {
        const matchedPath: RegExpMatchArray = currentPath
            ? currentPath.match(/([a-z\d\-]+)(\/*)$/i)
            : this.getCurrentPath().match(/([a-z\d\-]+)(\/*)$/i);
        let componentName: string = matchedPath !== null && matchedPath[1] ? matchedPath[1] : "";
        componentName = componentName.charAt(0).toUpperCase() + componentName.slice(1, componentName.length);
        componentName = componentName.replace(/-([a-z])/g, (match: string) => {
            return match[1].toUpperCase();
        });

        return componentName;
    }

    /**
     * Determine if the view is to the examples or detail of a component
     */
    private getComponentViewTypesByLocation(): ComponentViewTypes {
        return window
            && window.location
            && new RegExp(`${ComponentViewTypes[ComponentViewTypes.examples]}/?$`).test(window.location.pathname)
            ? ComponentViewTypes.examples
            : ComponentViewTypes.detail;
    }

    // private renderRoutes(): JSX.Element[] {
    //     return this.getRoutes(this.props.children, "/", SiteSlot.category)
    //         .map(this.renderComponentRoute);
    // }

    /**
     * Renders a route based on the active component
     */
    // private renderComponentRoute = (route: IComponentRoute | null): JSX.Element | null => {
    //     const path: string = route.route;

    //     return (
    //         <Route
    //             key={path}
    //             path={`${path}(|examples)`}
    //             exact={true}
    //         >
    //             {this.renderShellRow(route)}
    //         </Route>
    //     );
    // }

    // private renderShellRow(route: IComponentRoute): JSX.Element {
    //     const paneStyleSheet: Partial<ComponentStyles<IPaneClassNamesContract, IDevSiteDesignSystem>> = {
    //         pane: {
    //             backgroundColor: (config: IDevSiteDesignSystem): string => {
    //                 return config.lightGray;
    //             }
    //         }
    //     };

    //     return (
    //         <Row fill={true}>
    //             <Canvas>
    //                 <Row>
    //                     <ActionBar
    //                         onComponentViewChange={this.onComponentViewChange}
    //                         onFormToggle={this.onFormToggle}
    //                         onDevToolsToggle={this.onDevToolsToggle}
    //                         componentView={this.state.componentView}
    //                         formView={this.state.formView}
    //                         devToolsView={this.state.devToolsView}
    //                     />
    //                 </Row>
    //                 <div className={this.props.managedClasses.site_canvasContent}>
    //                     <ComponentView {...{ viewType: this.state.componentView }}>
    //                         {this.renderChildrenBySlot(this, ShellSlot.canvas)}
    //                         {this.renderComponentByRoute(route)}
    //                     </ComponentView>
    //                     <Row
    //                         resizable={true}
    //                         hidden={!this.state.devToolsView}
    //                         resizeFrom={RowResizeDirection.north}
    //                         minHeight={180}
    //                     >
    //                         {this.renderDevTools(route.schema)}
    //                     </Row>
    //                 </div>
    //             </Canvas>
    //             <Pane
    //                 hidden={!this.state.formView}
    //                 resizable={true}
    //                 resizeFrom={PaneResizeDirection.west}
    //                 jssStyleSheet={paneStyleSheet}
    //                 minWidth={324}
    //             >
    //                 {this.generateForm(route.schema)}
    //             </Pane>
    //         </Row>
    //     );
    // }

    private renderDevTools(schema: any): JSX.Element {
        if (this.state.devToolsView) {
            return (
                <DevTools
                    activeFramework={this.props.activeFramework}
                    activeSchema={schema}
                    activeComponentName={this.state.componentName}
                    activeFormData={this.getCurrentComponentData()}
                    childOptions={this.props.formChildOptions}
                    frameworks={Array.isArray(this.props.frameworks) ? this.props.frameworks : [this.props.frameworks]}
                    onToggleView={this.handleToggleDevToolsView}
                />
            );
        }
    }

    /**
     * Render the site title
     */
    private renderSiteTitle(): JSX.Element {
        const title: JSX.Element[] | undefined = this.renderChildrenBySlot(this, "title");

        return Array.isArray(title)
            ? (
                <span className={this.props.managedClasses.site_headerTitle}>
                    {title}
                </span>
            )
            : null;
    }

    private getCurrentComponentData(): any {
        if (this.state.componentView === ComponentViewTypes.examples) {
            return this.state.componentData[this.state.currentPath][this.state.activeComponentIndex];
        } else if (this.state.componentView === ComponentViewTypes.detail) {
            return this.state.detailViewComponentData[this.state.currentPath];
        }
    }

    private handleComponentDataChange = (data: any): void => {
        const currentPath: string = this.getCurrentPath();

        if (this.state.componentView === ComponentViewTypes.examples) {
            const componentData: IComponentData = Object.assign({}, this.state.componentData);
            componentData[currentPath][this.state.activeComponentIndex] = data;

            this.setState({
                componentData,
                currentPath
            });
        } else if (this.state.componentView === ComponentViewTypes.detail) {
            const detailViewComponentData: IComponentData = Object.assign({}, this.state.detailViewComponentData);
            detailViewComponentData[currentPath] = data;

            this.setState({
                detailViewComponentData,
                currentPath
            });
        }
    }

    private handleToggleDevToolsView = (): void => {
        this.setState({
            devToolsView: !this.state.devToolsView
        });
    }

    private get initialPath(): string {
        return Object.keys(this.getComponentData())[0];
    }

    private getCurrentPath = (): string => {
        return this.getComponentViewTypesByLocation() === ComponentViewTypes.detail
            ? window.location.pathname
            : window.location.pathname.slice(0, window.location.pathname.length - 9);
    }

    private generateForm(schema: any): JSX.Element {
            const componentData: any = this.state.componentView === ComponentViewTypes.examples
                ? Object.assign({}, this.state.componentData[this.state.currentPath][this.state.activeComponentIndex])
                : Object.assign({}, this.state.detailViewComponentData[this.state.currentPath]);

            return (
                <ConfigurationPanel
                    schema={schema}
                    data={componentData}
                    onChange={this.handleComponentDataChange}
                    formChildOptions={this.props.formChildOptions}
                />
            );
    }

    private generateComponentWrapperBackground(): string {
        return this.state.theme && this.state.theme.background ? this.state.theme.background : null;
    }

    private renderComponentByRoute(route: IComponentRoute): JSX.Element[] {
        if (route.exampleView && this.state.componentView === ComponentViewTypes.examples) {
            return route.exampleView.map((componentItem: JSX.Element, index: number) => {
                if (route.componentMapping && componentItem.props && componentItem.props.slot === ComponentViewSlot.example) {
                    return (
                        <ComponentWrapper
                            key={index}
                            onClick={this.handleComponentClick}
                            index={index}
                            background={this.generateComponentWrapperBackground()}
                            dir={isRTL(this.state.locale) ? Direction.rtl : Direction.ltr}
                            transparentBackground={this.state.componentBackgroundTransparent}
                            designSystem={componentItem.props.designSystem}
                            active={index === this.state.activeComponentIndex}
                            view={this.state.componentView}
                        >
                            <route.componentMapping {...this.state.componentData[route.route][index]} />
                        </ComponentWrapper>
                    );
                }
            });
        } else if (route.detailView && this.state.componentView === ComponentViewTypes.detail) {
            return route.detailView.map((componentItem: JSX.Element, index: number) => {
                if (route.componentMapping) {
                    return (
                        <React.Fragment key={index}>
                            {this.renderDetailViewComponent(componentItem, index, route)}
                            {this.renderDetailViewDocumentation(componentItem)}
                        </React.Fragment>
                    );
                }
            });
        }
    }

    private renderDetailViewComponent(component: JSX.Element, index: number, route: IComponentRoute): JSX.Element {
        if (component && component.props && component.props.slot === ComponentViewSlot.detailExample) {
            return (
                <ComponentWrapper
                    key={index}
                    index={index}
                    background={this.generateComponentWrapperBackground()}
                    dir={isRTL(this.state.locale) ? Direction.rtl : Direction.ltr}
                    transparentBackground={this.state.componentBackgroundTransparent}
                    designSystem={component.props.designSystem}
                    active={true}
                    view={this.state.componentView}
                >
                    <route.componentMapping {...this.state.detailViewComponentData[route.route]} />
                </ComponentWrapper>
            );
        }

        return null;
    }

    private renderDetailViewDocumentation(component: JSX.Element): JSX.Element {
        if (component && component.props && component.props.slot === ComponentViewSlot.detailDocumentation) {
            return component;
        }

        return null;
    }

    private handleComponentClick = (activeIndex: number): void => {
        this.setState({
            activeComponentIndex: activeIndex
        });
    }

    private onComponentViewChange = (type: ComponentViewTypes): void => {
        this.setState({
            componentView: type
        });
    }

    private onFormToggle = (): void => {
        this.setState({
            formView: !this.state.formView
        });
    }

    private onDevToolsToggle = (): void => {
        this.setState({
            devToolsView: !this.state.devToolsView
        });
    }

    private renderTransparencyToggle(): JSX.Element {
        if (this.props.showTransparencyToggle) {
            return (
                <button
                    aria-pressed={this.state.componentBackgroundTransparent}
                    onClick={this.handleTransparencyToggleClick}
                    className={this.props.managedClasses.site_transparencyToggleButton}
                >
                    <span dangerouslySetInnerHTML={{__html: glyphTransparency}}/>
                </button>
            );
        }
    }

    private renderStatus(): JSX.Element {
        if (this.props.showComponentStatus) {
            return (
                <React.Fragment>
                    <span className={this.getClassNames()}/>
                    {this.state.componentStatus}
                </React.Fragment>
            );
        }
    }

    private renderThemeSelect(): JSX.Element {
        /* tslint:disable-next-line */
        const className: string = `${this.props.managedClasses.site_infoBarConfiguration_base} ${this.props.managedClasses.site_infoBarConfiguration_theme}`

        if (Array.isArray(this.props.themes) && this.props.themes.length > 0) {
            return (
                <span className={className}>
                    <select
                        className={this.props.managedClasses.site_infoBarConfiguration_input}
                        onChange={this.handleThemeUpdate}
                        value={this.state.theme.displayName}
                    >
                        {this.renderThemes()}
                    </select>
                </span>
            );
        }
    }

    private getClassNames(): string {
        const classNames: string = this.props.managedClasses.site_statusIndicator;

        switch (this.state.componentStatus) {
            case Status.released:
                return `${classNames} ${this.props.managedClasses.site_statusReleased}`;
            case Status.alpha:
                return `${classNames} ${this.props.managedClasses.site_statusAlpha}`;
            default:
                return `${classNames} ${this.props.managedClasses.site_statusBeta}`;
        }
    }

    private renderLocales(): JSX.Element[] {
        return this.props.locales.map((locale: string, index: number): JSX.Element => {
            return (
                <option key={index}>
                    {locale}
                </option>
            );
        });
    }

    private renderThemes(): JSX.Element[] {
        return this.props.themes.map((theme: ITheme, index: number): JSX.Element => {
            return (
                <option key={theme.id} value={theme.id}>
                    {theme.displayName}
                </option>
            );
        });
    }

    private handleLocaleUpdate = (e: React.ChangeEvent<HTMLSelectElement>): void => {
        if (this.props.onUpdateDirection) {
            this.props.onUpdateDirection(isRTL(e.target.value) ? Direction.rtl : Direction.ltr);
        }

        this.setState({
            locale: e.target.value
        });
    }

    private handleThemeUpdate = (e: React.ChangeEvent<HTMLSelectElement>): void => {
        if (this.props.onUpdateTheme) {
            this.props.onUpdateTheme(e.target.value);
        }

        if (!this.props.activeTheme) {
            const selectedTheme: ITheme = this.props.themes.find((item: ITheme): boolean => {
                return item.id === e.target.value;
            });

            this.setState({
                theme: selectedTheme
            });
        }
    }

    /**
     * Gets all of the potential routes to build the shell
     */
    private getRoutes(
        children: React.ReactNode | React.ReactNodeArray,
        baseRoute: string,
        slot: SiteSlot,
        routes?: IComponentRoute[]
    ): IComponentRoute[] {
        let currentRoutes: IComponentRoute[] = routes;

        React.Children.forEach(children, (item: JSX.Element): void => {
            if (item && item.props && item.props.slot === slot) {
                currentRoutes = this.getCurrentRoute(item, slot, baseRoute, currentRoutes || []);
            }
        });

        return currentRoutes;
    }

    private getCurrentRoute(item: JSX.Element, slot: SiteSlot, baseRoute: string, currentRoutes: IComponentRoute[]): IComponentRoute[] {
        const currentRoute: IComponentRoute[] = currentRoutes;
        const itemRoute: string = `${baseRoute}${item.props.name}/`;
        const detailDocumentationSlotItems: JSX.Element[] = this.renderChildrenBySlot(item, ComponentViewSlot.detailDocumentation);
        const detailExampleSlotItems: JSX.Element[] = this.renderChildrenBySlot(item, ComponentViewSlot.detailExample);
        const detailSlotItems: JSX.Element[] = detailExampleSlotItems.concat(detailDocumentationSlotItems);
        const exampleSlotItems: JSX.Element[] = this.renderChildrenBySlot(item, ComponentViewSlot.example);

        if (exampleSlotItems && exampleSlotItems.length > 0) {
            currentRoute.push({
                route: this.convertToHyphenated(itemRoute),
                exampleView: exampleSlotItems,
                detailView: detailSlotItems,
                schema: item.props.schema,
                componentMapping: item.props.component,
                status: item.props.status
            });
        }

        if (item.props.children) {
            return this.getRoutes(item.props.children, itemRoute, slot, currentRoute);
        }

        return currentRoute;
    }

    /**
     * retreives the first component route whos route matches the provied path
     */
    private routeByPath(path: string): IComponentRoute {
        return this.getRoutes(this.props.children, "/", SiteSlot.category).find((route: IComponentRoute) => route.route === path);
    }

    private handlePaneCollapse = (): void => {
        this.setState({
            tableOfContentsCollapsed: !this.state.tableOfContentsCollapsed
        });
    }

    private handleTransparencyToggleClick = (e: React.MouseEvent<HTMLButtonElement>): void => {
        this.setState({
            componentBackgroundTransparent: !this.state.componentBackgroundTransparent
        });
    }

    private renderChildrenBySlot(component: any, slot: string): JSX.Element[] {
        if (component && component.props) {
            return React.Children.map(component.props.children, (child: JSX.Element, index: number) => {
                if (child && child.props && child.props.slot === slot) {
                    return child;
                }
            });
        }
    }

    /**
     * Renders the interactive table-of-contents, which is the primary
     * form of naviation around the site
     */
    private renderToc(): JSX.Element[] {
        return this.getToc(this.props.children, "/");
    }

    /**
     * Creates a table-of-contents from a given starting point.
     * This is used to recursivly generate sub-items
     */
    private getToc(items: React.ReactNode | React.ReactNode[], itemsPath: string): JSX.Element[] {
        return React.Children.toArray(items)
            .filter((item: React.ReactNode) => {
                return get(item, "props.slot") === SiteSlot.category;
            })
            .map((item: React.ReactElement<any>, index: number) => {
                return this.renderTocItem(index, itemsPath, item);
            });
    }

    /**
     * Adjust URL of TocItem based on the current component-view
     */
    private formatTocItemPathWithComponentViewTypes(path: string): string {
        return this.state.componentView === ComponentViewTypes.examples
            ? `${path}${ComponentViewTypes[ComponentViewTypes.examples]}/`
            : path;
    }

    private renderTocItem(index: number, itemsPath: string, child: JSX.Element): JSX.Element {
        const tocItemPath: string = this.convertToHyphenated(`${itemsPath}${child.props.name}/`);
        const onClick: (e: React.MouseEvent<HTMLButtonElement>) => void =
            (e: React.MouseEvent<HTMLButtonElement>): void => {
                this.setState({
                    activeComponentIndex: 0,
                    componentName: this.getComponentName(tocItemPath),
                    componentStatus: this.getComponentStatus(tocItemPath),
                    currentPath: tocItemPath
                });
            };

        return has(child, "props.name")
            ? (
                <TocItem
                    key={index}
                    active={new RegExp(this.state.currentPath).test(tocItemPath)}
                    heading={has(child, "props.children")}
                    onClick={onClick}
                    controls={uniqueId(this.convertToHyphenated(child.props.name))}
                    to={this.hasCanvasContent(child) ? this.formatTocItemPathWithComponentViewTypes(tocItemPath) : null}
                >
                    {this.renderTocItemCategory(child.props.name, this.renderTocItemCategoryIcon(child))}
                    {this.renderTocItemMenu(child, tocItemPath)}
                </TocItem>
            )
            : null;
    }

    private renderTocItemMenu(item: JSX.Element, tocItemPath: string): JSX.Element[] {
        return item.props.children && get(item.props.children, "props.slot") !== SiteSlot.categoryIcon
            ? this.getToc(item.props.children, tocItemPath)
            : null;
    }

    private renderTocItemCategory(name: string | any, icon?: JSX.Element): JSX.Element {
        const renderLayout: JSX.Element = icon
            ? (
                <div className={this.props.managedClasses.site_paneToggleButtonIconLayout}>
                    <span className={this.props.managedClasses.site_paneToggleButtonIcon}>{icon}</span>
                </div>
            )
            : null;

        if (this.state.tableOfContentsCollapsed) {
            return renderLayout;
        }

        return (
            <span className={this.props.managedClasses.site_paneTocRow}>
                {renderLayout}
                {icon ? <span className={this.props.managedClasses.site_paneTocTitle}>{name}</span> : name}
            </span>
        );
    }

    private renderTocItemCategoryIcon(item: JSX.Element): JSX.Element {
        if (Array.isArray(item.props.children)) {
            item.props.children.forEach((childElement: JSX.Element) => {
                if (childElement && childElement.props && childElement.props.slot === SiteSlot.categoryIcon) {
                    return childElement;
                }
            });
        } else {
            if (item.props.children.props && item.props.children.props.slot === SiteSlot.categoryIcon) {
                return item;
            }
        }

        return null;
    }

    /**
     * Determines if any children of a provided Element have content
     * that should be displayed in the canvas
     */
    private hasCanvasContent(item: React.ReactElement<any>): boolean {
        return React.Children.toArray(item.props.children).findIndex((child: React.ReactElement<any>) => {
            const slot: string = get(child, "props.slot");
            return slot === ComponentViewSlot.example
                || slot === ComponentViewSlot.detailExample
                || slot === ComponentViewSlot.detailDocumentation;
        }) !== -1;
    }

    private convertToHyphenated(name: string): string {
        return name.toLowerCase().replace(/\s/g, "-");
    }
}

export default manageJss(styles)(Site);
export {
    SiteTitle,
    SiteTitleBrand,
    SiteMenu,
    SiteMenuItem,
    SiteCategory,
    SiteCategoryDocumentation,
    SiteCategoryIcon,
    SiteCategoryItem
};
