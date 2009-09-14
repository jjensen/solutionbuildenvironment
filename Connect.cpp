// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn.h"
#include "Connect.h"

extern CAddInModule _AtlModule;

CComPtr<EnvDTE::_DTE> g_pDTE;

struct EnvVar
{
	CString m_name;
	CString m_value;
};

typedef list<EnvVar> EnvironmentVariablesList;
EnvironmentVariablesList g_savedEnvironmentVariables;

bool ResolveFilename(CString& filename, const CString& solutionDir,
					const CString& solutionName, bool makeAbsolute)
{
	// Resolve all environment variables.
	int position = -1;
	while (true)
	{
		// Search for $ symbols, denoting an environment variable.
		position = filename.Find('$', position + 1);
		if (position == -1)
			break;

		// Okay, there is an environment variable in there... resolve it.
		if (filename[position + 1] == '(')
		{
			// Fix by Steve Balderrama and Boris Glick: "There is a bug if
			// you are trying to do replacements on a string that includes
			// ‘)’ characters other than for $(xxx) constructs.  Parens are
			// legal in directory names and some tools install themselves
			// that way."
			int lastpos = filename.Find(')', position + 1);
			if (lastpos == -1)
			{
				// They forgot a closing )
				return false;
			}

			CString envStr = filename.Mid(position + 2, lastpos - (position + 2));
			CString env;

			// See if we can resolve it.  If not, then exit.
			if (envStr.CompareNoCase("solutiondir") == 0)
				env = solutionDir;
			else if (envStr.CompareNoCase("solutionname") == 0)
				env = solutionName;
			else
				env.GetEnvironmentVariable(envStr);

			// Okay, rebuild the string.
			filename = filename.Left(position) + env +
						filename.Right(filename.GetLength() - lastpos - 1);
		}
	}

	// Resolve all registry entries.
	position = -1;
	while (true)
	{
		// Search for % symbols, denoting a registry entry.
		position = filename.Find('%', position + 1);
		if (position == -1)
			break;

		// Okay, there is an registry entry in there... resolve it.
		if (filename[position + 1] == '(')
		{
			// Fix by Steve Balderrama and Boris Glick: "There is a bug if
			// you are trying to do replacements on a string that includes
			// ‘)’ characters other than for $(xxx) constructs.  Parens are
			// legal in directory names and some tools install themselves
			// that way."
			int lastpos = filename.Find(')', position + 1);
			if (lastpos == -1)
			{
				// A closing ) is missing.
				return false;
			}

			// HKLM\Software\Alias|Wavefront\Maya\4.5\Setup\InstallPath\MAYA_INSTALL_LOCATION
			CString reg = filename.Mid(position + 2, lastpos - (position + 2));

			CString root = reg.Left(4);
			reg = reg.Mid(5);

			// Strip the key.
			int slashPos = reg.ReverseFind('\\');
			if (slashPos == -1)
				break;
			CString key = reg.Mid(slashPos + 1);
			reg = reg.Left(slashPos);

			HKEY parentKey;
			if (root == "HKLM")
				parentKey = HKEY_LOCAL_MACHINE;
			else if (root = "HKCU")
				parentKey = HKEY_CURRENT_USER;

			CRegKey regKey;
			if (regKey.Open(parentKey, reg) == ERROR_SUCCESS)
			{
				char value[1024];
				ULONG numChars = 1024;

				if (regKey.QueryStringValue(key, value, &numChars) == ERROR_SUCCESS)
				{
					// Okay, rebuild the string.
					filename = filename.Left(position) + value +
								filename.Right(filename.GetLength() - lastpos - 1);
				}
			}
		}
	}

	if (makeAbsolute)
	{
		// Make the intermediate directory an absolute path.
		char currentDirectory[MAX_PATH];
		GetCurrentDirectory(MAX_PATH, currentDirectory);

		SetCurrentDirectory(solutionDir);

		char path[MAX_PATH];
		_fullpath(path, filename, MAX_PATH);
		filename = path;

		SetCurrentDirectory(currentDirectory);
	}

	return true;
}


bool ParseSlnEnvFile(CString filename, const CString& solutionDir,
					 const CString& solutionName,
					 const CString& solutionConfigurationName)
{
	// Add .env.
	filename += ".slnenv";

	int lineNumber = 0;

	FILE* file = fopen(filename, "rt");
	if (file)
	{
		while (!feof(file))
		{
			CString line;
			LPSTR buf = line.GetBuffer(2048);
			fgets(buf, 2048, file);
			line.ReleaseBuffer();
			lineNumber++;

			// Get rid of newline and spaces at end of line
			// Get rid of spaces at beginning of line
			line.TrimRight(" \n");
			line.TrimLeft(' ');

			// Find hash marks. Remove all to the right of first one found.
			int commentIndex = line.Find('#');
			if (commentIndex != -1)
				line = line.Left(commentIndex);
			commentIndex = line.Find("--");
			if (commentIndex != -1)
				line = line.Left(commentIndex);
			commentIndex = line.Find("//");
			if (commentIndex != -1)
				line = line.Left(commentIndex);

			line.Trim();

			// Test for blank line. This should be done after hash mark find code
			if (line.IsEmpty())
				continue;

			int equalsPos = line.Find('=');

			// This block checks for include and forceinclude
			if (equalsPos == -1)
			{
				int spacePos = line.Find(' ');
				if (spacePos != -1)
				{
					CString command = line.Left(spacePos);
					if (command == "include"  ||  command == "forceinclude")
					{
						CString newFilename = line.Mid(spacePos + 1);
						if (!ResolveFilename(newFilename, solutionDir, solutionName, false))
						{
							fclose(file);
							CString err;
							err.Format("Line #%d of file [%s] is invalid.  Perhaps a closing ) is missing?.", lineNumber, filename);
							throw err;
						}

						if (!ParseSlnEnvFile(newFilename, solutionDir, solutionName, solutionConfigurationName))
						{
							if (command == "forceinclude")
							{
								fclose(file);
								CString err;
								err.Format("Line #%d of file [%s] is invalid.  include [%s.slnenv] not found.", lineNumber, filename, newFilename);
								throw err;
							}
						}
						continue;
					}
					else
					{
						fclose(file);
						CString err;
						err.Format("Line #%d of file [%s] is invalid.  include Filename or name=value expected.", lineNumber, filename);
						throw err;
					}
				}
				else
				{
					fclose(file);
					CString err;
					err.Format("Line #%d of file [%s] is invalid.  include Filename or name=value expected.", lineNumber, filename);
					throw err;
				}
			}

			// See if it is ?=
			bool isQuestion = false;
			int extraEquals = 0;
			if (equalsPos > 0)
			{
				isQuestion = line[equalsPos - 1] == '?';
				if (isQuestion)
				{
					equalsPos--;
					extraEquals = 1;
				}
			}

			// Figure out the Key
			CString name = line.Left(equalsPos);
			name.Trim();

			bool makeAbsolute = false;
			if (name[0] == '!')
			{
				makeAbsolute = true;
				name = name.Mid(1);
			}

			// The key is configuration specific if ':' is used
			int colonPos = name.Find(':');
			if (colonPos != -1)
			{
				CString configName = name.Left(colonPos);
				if (configName != solutionConfigurationName)
					continue;
				name = name.Mid(colonPos + 1);
			}

			// Figure out the value
			CString value = line.Mid(equalsPos + 1 + extraEquals);
			value.Trim();

			if (!ResolveFilename(value, solutionDir, solutionName, makeAbsolute))
			{
				fclose(file);
				CString err;
				err.Format("Line #%d of file [%s] is invalid.  Perhaps a closing ) is missing?.", lineNumber, filename);
				throw err;
			}

			EnvVar envVar;
			envVar.m_name = name;
			envVar.m_value.GetEnvironmentVariable(name);
			if (!isQuestion  ||  envVar.m_value.IsEmpty())
			{
				g_savedEnvironmentVariables.push_back(envVar);

				SetEnvironmentVariable(name, value);
			}
		}

		fclose(file);
		return true;
	}

	return false;
}


void UndoEnvironmentChanges()
{
	for (EnvironmentVariablesList::iterator it = g_savedEnvironmentVariables.begin();
		it != g_savedEnvironmentVariables.end(); ++it)
	{
		EnvVar& envVar = (*it);
		SetEnvironmentVariable(envVar.m_name, envVar.m_value);
	}
}


void ReadSolutionEnvironment()
{
	try
	{
		UndoEnvironmentChanges();
		g_savedEnvironmentVariables.clear();

		CComPtr<EnvDTE::_Solution> pSolution;
		g_pDTE->get_Solution(&pSolution);
		if (!pSolution)
			return;

		CComBSTR bstrFullPath;
		pSolution->get_FullName(&bstrFullPath);

		// Strip the .sln.
		CString fullPath(bstrFullPath);
		int dotPos = fullPath.ReverseFind('.');
		if (dotPos == -1)
			return;
		fullPath = fullPath.Left(dotPos);

		CString solutionDir = fullPath;
		int slashPos = solutionDir.ReverseFind('\\');
		if (slashPos == -1)
			return;
		solutionDir = solutionDir.Left(slashPos);

		CString solutionName = fullPath;
		solutionName = solutionName.Mid(slashPos + 1);

		CComPtr<EnvDTE::SolutionBuild> pSolutionBuild;
		pSolution->get_SolutionBuild(&pSolutionBuild);
		if (!pSolutionBuild)
			return;

		CComBSTR bstrConfig;

		CComPtr<EnvDTE::SolutionConfiguration> pSolutionConfiguration;
		pSolutionBuild->get_ActiveConfiguration(&pSolutionConfiguration);
		if (pSolutionConfiguration)
		{
			pSolutionConfiguration->get_Name(&bstrConfig);
		}

		CString solutionConfigurationName(bstrConfig);

		ParseSlnEnvFile(fullPath, solutionDir, solutionName, solutionConfigurationName);
	}
	catch (CString& str)
	{
		::MessageBox(NULL, str, "Error", MB_OK);
	}
	catch (...)
	{
		::MessageBox(NULL, "The Solution Build Environment add-in crashed.", "Error", MB_OK);
	}
}


class SolutionEventsSink70 : public IDispEventImpl<1, SolutionEventsSink70, &__uuidof(EnvDTE::_dispSolutionEvents), &EnvDTE::LIBID_EnvDTE, 7, 0>
{
public:
	BEGIN_SINK_MAP(SolutionEventsSink70)
		SINK_ENTRY_EX(1, __uuidof(EnvDTE::_dispSolutionEvents), 1, Opened)
	END_SINK_MAP()

	void __stdcall Opened()
	{
		ReadSolutionEnvironment();
	}
};


class SolutionEventsSink80 : public IDispEventImpl<1, SolutionEventsSink80, &__uuidof(EnvDTE::_dispSolutionEvents), &EnvDTE::LIBID_EnvDTE, 8, 0>
{
public:
	BEGIN_SINK_MAP(SolutionEventsSink80)
		SINK_ENTRY_EX(1, __uuidof(EnvDTE::_dispSolutionEvents), 1, Opened)
	END_SINK_MAP()

	void __stdcall Opened()
	{
		ReadSolutionEnvironment();
	}
};


class SolutionEventsSink90 : public IDispEventImpl<1, SolutionEventsSink90, &__uuidof(EnvDTE::_dispSolutionEvents), &EnvDTE::LIBID_EnvDTE, 9, 0>
{
public:
	BEGIN_SINK_MAP(SolutionEventsSink90)
		SINK_ENTRY_EX(1, __uuidof(EnvDTE::_dispSolutionEvents), 1, Opened)
	END_SINK_MAP()

	void __stdcall Opened()
	{
		ReadSolutionEnvironment();
	}
};


SolutionEventsSink70* m_SolutionEventsSink70;
SolutionEventsSink80* m_SolutionEventsSink80;
SolutionEventsSink90* m_SolutionEventsSink90;
CComPtr<EnvDTE::_SolutionEvents> pSolutionEvents;


class BuildEventsSink70 : public IDispEventImpl<1, BuildEventsSink70, &__uuidof(EnvDTE::_dispBuildEvents), &EnvDTE::LIBID_EnvDTE, 7, 0>
{
public:
	BEGIN_SINK_MAP(BuildEventsSink70)
		SINK_ENTRY_EX(1, __uuidof(EnvDTE::_dispBuildEvents), 3, OnBuildBegin)
	END_SINK_MAP()

    void __stdcall OnBuildBegin(
					EnvDTE::vsBuildScope Scope,
                    EnvDTE::vsBuildAction Action)
	{
		// Marc Roth: Added vsBuildActionClean.
		if (Action == EnvDTE::vsBuildActionBuild ||
			Action == EnvDTE::vsBuildActionRebuildAll ||
			Action == EnvDTE::vsBuildActionClean)
		{
			ReadSolutionEnvironment();
		}
	}
};


class BuildEventsSink80 : public IDispEventImpl<1, BuildEventsSink80, &__uuidof(EnvDTE::_dispBuildEvents), &EnvDTE::LIBID_EnvDTE, 8, 0>
{
public:
	BEGIN_SINK_MAP(BuildEventsSink80)
		SINK_ENTRY_EX(1, __uuidof(EnvDTE::_dispBuildEvents), 3, OnBuildBegin)
	END_SINK_MAP()

    void __stdcall OnBuildBegin(
					EnvDTE::vsBuildScope Scope,
                    EnvDTE::vsBuildAction Action)
	{
		// Marc Roth: Added vsBuildActionClean.
		if (Action == EnvDTE::vsBuildActionBuild ||
			Action == EnvDTE::vsBuildActionRebuildAll ||
			Action == EnvDTE::vsBuildActionClean)
		{
			ReadSolutionEnvironment();
		}
	}
};


class BuildEventsSink90 : public IDispEventImpl<1, BuildEventsSink90, &__uuidof(EnvDTE::_dispBuildEvents), &EnvDTE::LIBID_EnvDTE, 9, 0>
{
public:
	BEGIN_SINK_MAP(BuildEventsSink90)
		SINK_ENTRY_EX(1, __uuidof(EnvDTE::_dispBuildEvents), 3, OnBuildBegin)
	END_SINK_MAP()

    void __stdcall OnBuildBegin(
					EnvDTE::vsBuildScope Scope,
                    EnvDTE::vsBuildAction Action)
	{
		// Marc Roth: Added vsBuildActionClean.
		if (Action == EnvDTE::vsBuildActionBuild ||
			Action == EnvDTE::vsBuildActionRebuildAll ||
			Action == EnvDTE::vsBuildActionClean)
		{
			ReadSolutionEnvironment();
		}
	}
};


BuildEventsSink70* m_BuildEventsSink70;
BuildEventsSink80* m_BuildEventsSink80;
BuildEventsSink90* m_BuildEventsSink90;
CComPtr<EnvDTE::_BuildEvents> pBuildEvents;

CComBSTR version;

void CConnect::FreeEvents()
{
	if ( pSolutionEvents )
	{
		if (version == "7.00"  ||  version == "7.10")
		{
			m_SolutionEventsSink70->DispEventUnadvise((IUnknown*)pSolutionEvents.p);
			delete m_SolutionEventsSink70;
			m_SolutionEventsSink70 = NULL;
		}
		else if (version == "8.0")
		{
			m_SolutionEventsSink80->DispEventUnadvise((IUnknown*)pSolutionEvents.p);
			delete m_SolutionEventsSink80;
			m_SolutionEventsSink80 = NULL;
		}
		else if (version == "9.0")
		{
			m_SolutionEventsSink80->DispEventUnadvise((IUnknown*)pSolutionEvents.p);
			delete m_SolutionEventsSink80;
			m_SolutionEventsSink80 = NULL;
		}
		pSolutionEvents = NULL;
	}

	if ( pBuildEvents )
	{
		if (version == "7.00"  ||  version == "7.10")
		{
			m_BuildEventsSink70->DispEventUnadvise((IUnknown*)pBuildEvents.p);
			delete m_BuildEventsSink70;
			m_BuildEventsSink70 = NULL;
		}
		else if (version == "8.0")
		{
			m_BuildEventsSink80->DispEventUnadvise((IUnknown*)pBuildEvents.p);
			delete m_BuildEventsSink80;
			m_BuildEventsSink80 = NULL;
		}
		else if (version == "9.0")
		{
			m_BuildEventsSink80->DispEventUnadvise((IUnknown*)pBuildEvents.p);
			delete m_BuildEventsSink80;
			m_BuildEventsSink80 = NULL;
		}
		pBuildEvents = NULL;
	}
}


// CConnect
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	// Extra error checking, because I still don't understand the source of
	// the startup and shutdown crashes.
	g_pDTE = NULL;
	m_pAddInInstance = NULL;

	CComPtr<EnvDTE::Events> pEvents;

	HRESULT hr = S_OK;
	pApplication->QueryInterface(__uuidof(EnvDTE::_DTE), (LPVOID*)&g_pDTE);
	pAddInInst->QueryInterface(__uuidof(EnvDTE::AddIn), (LPVOID*)&m_pAddInInstance);

	g_pDTE->get_Version(&version);

	IfFailGo(g_pDTE->get_Events(&pEvents));

	FreeEvents();

	if(SUCCEEDED(pEvents->get_SolutionEvents((EnvDTE::_SolutionEvents**)&pSolutionEvents)))
	{
		if (version == "7.00"  ||  version == "7.10")
		{
			m_SolutionEventsSink70 = new SolutionEventsSink70;
			m_SolutionEventsSink70->DispEventAdvise((IUnknown*)pSolutionEvents.p);
		}
		else if (version == "8.0")
		{
			m_SolutionEventsSink80 = new SolutionEventsSink80;
			m_SolutionEventsSink80->DispEventAdvise((IUnknown*)pSolutionEvents.p);
		}
		else if (version == "9.0")
		{
			m_SolutionEventsSink80 = new SolutionEventsSink80;
			m_SolutionEventsSink80->DispEventAdvise((IUnknown*)pSolutionEvents.p);
		}
	}

	if(SUCCEEDED(pEvents->get_BuildEvents((EnvDTE::_BuildEvents**)&pBuildEvents)))
	{
		if (version == "7.00"  ||  version == "7.10")
		{
			m_BuildEventsSink70 = new BuildEventsSink70;
			m_BuildEventsSink70->DispEventAdvise((IUnknown*)pBuildEvents.p);
		}
		else if (version == "8.0")
		{
			m_BuildEventsSink80 = new BuildEventsSink80;
			m_BuildEventsSink80->DispEventAdvise((IUnknown*)pBuildEvents.p);
		}
		else if (version == "9.0")
		{
			m_BuildEventsSink80 = new BuildEventsSink80;
			m_BuildEventsSink80->DispEventAdvise((IUnknown*)pBuildEvents.p);
		}
	}


Error:
	return hr;
}

STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	FreeEvents();

	g_pDTE = NULL;
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

