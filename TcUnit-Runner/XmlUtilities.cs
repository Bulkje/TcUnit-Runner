﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TcUnit.TcUnit_Runner
{
    class XmlUtilities
    {
        // Singleton constructor
        private XmlUtilities()
        { }

        /// <summary>
        /// Sets the <Disabled> and <AutoStart>-tags
        /// </summary>
        /// <param name="rtXml">The XML-string coming from the real-time task automation interface object</param>
        /// <returns>String with parameters updated</returns>
        public static string SetDisabledAndAndAutoStartOfRealTimeTaskXml(string rtXml, bool disabled, bool autostart)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rtXml);
            XmlNode disabledNode = xmlDoc.SelectSingleNode("/TreeItem/" + "Disabled");
            if (disabledNode != null)
            {
                disabledNode.InnerText = disabled.ToString().ToLower();
            }
            else
            {
                return "";
            }

            XmlNode autoStartNode = xmlDoc.SelectSingleNode("/TreeItem/TaskDef/" + "AutoStart");
            if (autoStartNode != null)
            {
                autoStartNode.InnerText = autostart.ToString().ToLower();
            }
            else
            {
                return "";
            }
            return xmlDoc.OuterXml;
        }

        /// <summary>
        /// Gets the <ItemName> from the XML
        /// </summary>
        public static string GetItemNameFromRealTimeTaskXML(string rtXml) {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rtXml);
            XmlNode itemNameNode = xmlDoc.SelectSingleNode("/TreeItem/" + "ItemName");
            if (itemNameNode != null)
            {
                return itemNameNode.InnerText;
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Check whether a TwinCAT project is pinned
        /// </summary>
        /// <param name="TwinCATProjectFilePath">The complete file path to the TwinCAT (*.tsproj)-file</param>
        /// <returns>True if pinned, otherwise false</returns>
        public static bool IsTwinCATProjectPinned(string TwinCATProjectFilePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(TwinCATProjectFilePath);

            XmlNode nodePinnedVersion = xmlDoc.SelectSingleNode("/TcSmProject");
            var attrPinnedVersion = nodePinnedVersion.Attributes["TcVersionFixed"];

            if (nodePinnedVersion == null || attrPinnedVersion == null)
            {
                return false;
            }
            return Convert.ToBoolean(attrPinnedVersion.InnerText);
        }

        /// <summary>
        /// Returns the Port that is used for a PLC node
        /// </summary>
        public static int AmsPort(string plcXml)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(plcXml);

            XmlNode nodeAdsPort = xmlDoc.SelectSingleNode("/TreeItem/PlcProjectDef/" + "AdsPort");
            return Convert.ToInt32(nodeAdsPort.InnerText);
        }
        
        /// <summary>
        /// Returns the whole XML file as string
        /// </summary>
        public static string XMLstring(string path)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(path);
            return xmlDoc.OuterXml;
        }

        public static string addCompilerDefine(string plcproj, string compilerdefine)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(plcproj);
            XmlNode target = xmlDoc.SelectSingleNode("/TreeItem/PlcProjectDef");

            //Create new element for compilerdefine
            XmlElement elem = xmlDoc.CreateElement("CompilerDefines");
            elem.InnerText = compilerdefine;

            //add the node to the document
            //if node exists it is removed from its original position and added to its target position.
            target.AppendChild(elem);

            return xmlDoc.OuterXml;
        }
    }
}
