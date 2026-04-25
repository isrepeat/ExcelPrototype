
/*
// DevUI.xml

<page>
    <grid>
        <control>
        </control>
    </grid>
</page>
*/



class XmlDom;
class XmlNode;
class XmlLayoutEngine;

class Page {
public:
    Page(const std::string& xmlFilePath) {
        xmlDom.loadFromFile(xmlFilePath);
    }
    void Render() {
         // reset layout state ...
        XmlLayoutEngine::RenderNode(xmlDom.getRootElement(), this); // render pageNode
        // applying styles ...
    }

private:
    XmlDom xmlDom;
};


XmlLayoutEngine::RenderNode(XmlNode xmlNode, Page* page) {
    // render node
    switch (xmlNode.getType()) {
        case XmlNodeType::Page:
            PageRenderer::Render(xmlNode, page);
            break;
        case XmlNodeType::Grid:
            // render grid
            break;
        case XmlNodeType::StackPanel:
            // render stack panel
            break;
        case XmlNodeType::Control:
            ControlRenderer::Render(xmlNode, page);
            break;
    }
}

class PageRenderer {
public:
    static void Render(XmlNode xmlPageNode, Page* page) {
        for (auto& child : xmlPageNode.children()) {
            // render child nodes (grid, stackpanel, control)
            XmlLayoutEngine::RenderNode(child, page);
        }
    }
};

class ControlRenderer {
public:
    static void Render(XmlNode xmlControlNode, Page* page) {
        switch (xmlControlNode.getAttribute("type")) {
            case ControlType::Button:
                // render button
                break;
            case ControlType::Select:
                // render select
                break;
        }
    }
};