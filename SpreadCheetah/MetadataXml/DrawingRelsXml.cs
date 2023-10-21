using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingRelsXml
{
    // TODO: Example with a single image in sheet 1 (drawing1.xml.rels):
    //<?xml version = "1.0" encoding="UTF-8"?>
    //<Relationships xmlns = "http://schemas.openxmlformats.org/package/2006/relationships" >
    //    < Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
    //</Relationships>


    // TODO: Example with two images in sheet 1 (drawing1.xml.rels):
    //<?xml version = "1.0" encoding="UTF-8"?>
    //<Relationships xmlns = "http://schemas.openxmlformats.org/package/2006/relationships" >
    //    < Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
    //    <Relationship Id = "rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
    //</Relationships>


    // TODO: Example with two images in two sheets:
    // drawing1.xml.rels:
    //<?xml version = "1.0" encoding="UTF-8"?>
    //<Relationships xmlns = "http://schemas.openxmlformats.org/package/2006/relationships" >
    //    < Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image3.jpeg"/>
    //    <Relationship Id = "rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image4.png"/>
    //</Relationships>

    // drawing2.xml.rels
    //<?xml version="1.0" encoding="UTF-8"?>
    //<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    //    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image5.png"/>
    //    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image6.png"/>
    //</Relationships>
}
