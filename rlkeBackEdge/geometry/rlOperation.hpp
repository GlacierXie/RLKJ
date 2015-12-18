#pragma once

class rlOperation
{
public:
	void start()
	{
		will_begin();
		beginning();
		ended();
	}

protected:
	virtual void will_begin(){}
	virtual void beginning(){}
	virtual void ended(){}
};

