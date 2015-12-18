#include "stdafx.h"
#include "rlHalfEdge2d.h"


rlHalfEdge2d::rlHalfEdge2d()
	: _is_positive(true)
	, _owner(0)
{
}


rlHalfEdge2d::~rlHalfEdge2d()
{
}

void rlHalfEdge2d::set_owner(rlId edge)
{
	_owner = edge;
}

rlId rlHalfEdge2d::owner() const
{
	return _owner;
}

void rlHalfEdge2d::set_is_positive(bool value)
{
	_is_positive = value;
}

bool rlHalfEdge2d::is_positive() const
{
	return _is_positive;
}
