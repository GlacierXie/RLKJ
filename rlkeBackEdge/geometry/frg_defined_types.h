#pragma once
#include <boost/property_tree/ptree.hpp>
#include <boost/property_tree/xml_parser.hpp> 


typedef boost::property_tree::ptree rlXtree;
class frgGlobals
{
public:
	static AcGeTol Gtol;

	static const rlXtree &code_rule_xtree()
	{
		if (_code_rule_ptree.empty())
		{
			boost::property_tree::read_xml("code_rule.xml", _code_rule_ptree,
				boost::property_tree::xml_parser::trim_whitespace);
		}
		return _code_rule_ptree;
	}

private:
	static rlXtree _code_rule_ptree;
};