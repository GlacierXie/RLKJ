#include "stdafx.h"
#include "rlVertex2d.h"
#include "rlEdge2d.h"


rlVertex2d::rlVertex2d()
	: _x(0)
	, _y(0)
{
}

rlVertex2d::~rlVertex2d()
{
}

void rlVertex2d::set_x(double x)
{
	_x = x;
}

double rlVertex2d::x() const
{
	return _x;
}

void rlVertex2d::set_y(double y)
{
	_y = y;
}

double rlVertex2d::y() const
{
	return _y;
}

void rlVertex2d::insert(rlId edge)
{
	_edges.insert(edge);
}

void rlVertex2d::erase(rlId edge)
{
	_edges.erase(edge);
}

const std::set<rlId> &rlVertex2d::edges() const
{
	return _edges;
}

bool rlVertex2d::find(rlId edge) const
{
	if (_edges.find(edge) != _edges.end())
		return true;
	return false;
}

bool rlVertex2d::has_coedge(const rlVertex2d *vertex) const
{
	std::set<rlId>::const_iterator it = _edges.begin();
	for (; it != _edges.end(); ++it)
	{
		if (vertex->find(*it))
			return true;
	}
	return false;
}

bool rlVertex2d::has_another_end(rlId &vertex, const rlEdge2d *edge) const
{
	
	if (edge->v1() == id())
		vertex = edge->v2();
	else if (edge->v2() == id())
		vertex = edge->v1();

	if (vertex != 0)
		return true;
	return false;
}
