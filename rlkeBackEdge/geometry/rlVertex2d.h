#pragma once
#include "rlTopologicalEntity2d.h"


class rlEdge2d;
class rlVertex2d :
	public rlTopologicalEntity2d
{
	friend class rlTopologicalEntSet;
public:
	rlSubClassTypeDef(rlTopologicalEntity2d,emVertex2d)

	void set_x(double x);
	double x() const;

	void set_y(double y);
	double y() const;

	void insert(rlId edge);
	void erase(rlId edge);
	const std::set<rlId> &edges() const;

	bool find(rlId edge) const;
	bool has_coedge(const rlVertex2d *vertex) const;
	bool has_another_end(rlId &vertex, const rlEdge2d *edge) const;

private:
	rlVertex2d();
	~rlVertex2d();

private:
	double _x;
	double _y;
	std::set<rlId> _edges;
};

