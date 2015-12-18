#include "stdafx.h"
#include "rlEdge2d.h"


rlEdge2d::rlEdge2d()
	: _v1(0)
	, _v2(0)
	, _he12(0)
	, _he21(0)
{
}

rlEdge2d::~rlEdge2d()
{
}

void rlEdge2d::set_v1(rlId v)
{
	_v1 = v;
}

rlId rlEdge2d::v1() const
{
	return _v1;
}

void rlEdge2d::set_v2(rlId v)
{
	_v2 = v;
}

rlId rlEdge2d::v2() const
{
	return _v2;
}

void rlEdge2d::set_he12(rlId value)
{
	_he12 = value;
}

rlId rlEdge2d::he12() const
{
	return _he12;
}

void rlEdge2d::set_he21(rlId value)
{
	_he21 = value;
}

rlId rlEdge2d::he21() const
{
	return _he21;
}
